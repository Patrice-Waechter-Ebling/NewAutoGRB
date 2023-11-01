VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmProjSoumMec 
   BackColor       =   &H00000000&
   Caption         =   "Projets / Soumissions Mécaniques"
   ClientHeight    =   7770
   ClientLeft      =   225
   ClientTop       =   645
   ClientWidth     =   13380
   Icon            =   "frmProjSoumMec.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmProjSoumMec.frx":2CFA
   ScaleHeight     =   7770
   ScaleWidth      =   13380
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbOuvertFerme 
      Height          =   315
      ItemData        =   "frmProjSoumMec.frx":206AC
      Left            =   4560
      List            =   "frmProjSoumMec.frx":206B6
      Style           =   2  'Dropdown List
      TabIndex        =   122
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame fraFournisseur 
      BackColor       =   &H00000000&
      Caption         =   "Fournisseurs"
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
      Height          =   2415
      Left            =   1320
      TabIndex        =   14
      Top             =   1800
      Visible         =   0   'False
      Width           =   11415
      Begin VB.CommandButton cmdSupprimerFRS 
         Caption         =   "Supprimer"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton cmdAnnulerFRS 
         Caption         =   "Annuler"
         Height          =   375
         Left            =   9000
         TabIndex        =   17
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdOKFRS 
         Caption         =   "OK"
         Height          =   375
         Left            =   10200
         TabIndex        =   18
         Top             =   1920
         Width           =   1095
      End
      Begin MSComctlLib.ListView lvwFournisseur 
         Height          =   1575
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   2778
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fournisseur"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Pers. Ress."
            Object.Width           =   1984
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   1746
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Par"
            Object.Width           =   805
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Valide"
            Object.Width           =   1746
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Prix listé"
            Object.Width           =   1614
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Escompte"
            Object.Width           =   1561
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Prix net"
            Object.Width           =   1614
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Prix spécial"
            Object.Width           =   1720
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Quoter"
            Object.Width           =   1191
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Stock"
            Object.Width           =   1499
         EndProperty
      End
   End
   Begin VB.Frame fraPieceTrouve 
      BackColor       =   &H00000000&
      Caption         =   "Pièces trouvées"
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
      Height          =   2775
      Left            =   2160
      TabIndex        =   19
      Top             =   2280
      Visible         =   0   'False
      Width           =   10335
      Begin VB.CommandButton cmdOKPieceTrouve 
         Caption         =   "OK"
         Height          =   375
         Left            =   9120
         TabIndex        =   22
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdAnnulerPieceTrouve 
         Caption         =   "Annuler"
         Height          =   375
         Left            =   7920
         TabIndex        =   21
         Top             =   2280
         Width           =   1095
      End
      Begin MSComctlLib.ListView lvwPieceTrouve 
         Height          =   1935
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   3413
         SortKey         =   1
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "PIECE_GRB"
            Object.Width           =   2408
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "No. d'item"
            Object.Width           =   3254
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Catégorie"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Manufacturier"
            Object.Width           =   2037
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Description française"
            Object.Width           =   7144
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Description anglaise"
            Object.Width           =   7144
         EndProperty
      End
   End
   Begin MSComCtl2.MonthView mvwDateFacturation 
      Height          =   2370
      Left            =   1080
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   0
      Appearance      =   1
      StartOfWeek     =   10682369
      CurrentDate     =   38310
   End
   Begin MSComctlLib.ListView lvwBavard 
      Height          =   1575
      Left            =   120
      TabIndex        =   24
      Top             =   840
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   2778
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nom de l'employé"
         Object.Width           =   3201
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   1746
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Heure"
         Object.Width           =   1931
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Qté"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "No. Item"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvwHistorique 
      Height          =   1575
      Left            =   120
      TabIndex        =   23
      Top             =   840
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   2778
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nom de l'employé"
         Object.Width           =   3201
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   1746
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Heure"
         Object.Width           =   1931
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Valeur"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame fraDateRequise 
      BackColor       =   &H00000000&
      Caption         =   "Date Requise"
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
      Height          =   2895
      Left            =   3480
      TabIndex        =   95
      Top             =   4680
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton cmdOKDateRequise 
         Caption         =   "OK"
         Height          =   375
         Left            =   3480
         TabIndex        =   97
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdAnnulerDateRequise 
         Caption         =   "Annuler"
         Height          =   375
         Left            =   3480
         TabIndex        =   98
         Top             =   1200
         Width           =   1095
      End
      Begin MSComCtl2.MonthView mvwDateRequise 
         Height          =   2370
         Left            =   600
         TabIndex        =   96
         Top             =   360
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   0
         Appearance      =   1
         StartOfWeek     =   10682369
         CurrentDate     =   38247
      End
   End
   Begin VB.Frame fraCommentaire 
      BackColor       =   &H00000000&
      Caption         =   "Commentaire"
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
      Height          =   2895
      Left            =   3480
      TabIndex        =   91
      Top             =   4680
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton cmdAnnulerCommentaire 
         Caption         =   "Annuler"
         Height          =   375
         Left            =   3600
         TabIndex        =   94
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdOKCommentaire 
         Caption         =   "OK"
         Height          =   375
         Left            =   3600
         TabIndex        =   93
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtCommentaire 
         Height          =   2415
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   92
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.TextBox txtPrixSoumission 
      Height          =   285
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   61
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdRetour 
      Caption         =   "Retour"
      Height          =   375
      Left            =   6480
      TabIndex        =   114
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdSortieMagasin 
      Caption         =   "Magasin"
      Height          =   375
      Left            =   5400
      TabIndex        =   107
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   3360
      TabIndex        =   105
      Top             =   7320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdRapportFACT 
      Caption         =   "Fact"
      Height          =   375
      Left            =   2280
      TabIndex        =   102
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdMauvaisPrix 
      Caption         =   "Prix"
      Height          =   375
      Left            =   3360
      TabIndex        =   103
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdMaterielInutile 
      Caption         =   "Inutile"
      Height          =   375
      Left            =   5400
      TabIndex        =   108
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdCatalogue 
      Caption         =   "Catalogue"
      Height          =   375
      Left            =   6480
      TabIndex        =   112
      Top             =   7320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdExtra 
      Caption         =   "Extra"
      Height          =   375
      Left            =   6480
      TabIndex        =   113
      Top             =   7320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdDemande 
      Caption         =   "Demande"
      Height          =   375
      Left            =   5400
      TabIndex        =   106
      Top             =   7320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdAnglaisFrancais 
      Caption         =   "Anglais"
      Height          =   375
      Left            =   1200
      TabIndex        =   101
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdBonCommande 
      Caption         =   "Bon Com."
      Height          =   375
      Left            =   5400
      TabIndex        =   109
      Top             =   7320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdCopier 
      Caption         =   "Copier"
      Height          =   375
      Left            =   5400
      TabIndex        =   111
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdCreerProjet 
      Caption         =   "Créer proj."
      Height          =   375
      Left            =   5400
      TabIndex        =   110
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdSupprimer 
      Caption         =   "Supprimer"
      Height          =   375
      Left            =   8640
      TabIndex        =   117
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdModifier 
      Caption         =   "Modifier"
      Height          =   375
      Left            =   9720
      TabIndex        =   119
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdAjouter 
      Caption         =   "Ajouter"
      Height          =   375
      Left            =   7560
      TabIndex        =   115
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdFermer 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   10800
      TabIndex        =   120
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdAnnuler 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   9720
      TabIndex        =   118
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdEnregistrer 
      Caption         =   "Enregistrer"
      Height          =   375
      Left            =   8640
      TabIndex        =   116
      Top             =   7320
      Width           =   975
   End
   Begin VB.Frame fraPrix 
      BackColor       =   &H00000000&
      Caption         =   "Fournisseurs"
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
      Height          =   2295
      Left            =   720
      TabIndex        =   75
      Top             =   4680
      Visible         =   0   'False
      Width           =   8895
      Begin VB.TextBox txtPrixSpecial 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4920
         TabIndex        =   85
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton cmdAnnulerPrix 
         Caption         =   "Annuler"
         Height          =   375
         Left            =   6240
         TabIndex        =   89
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdOKPrix 
         Caption         =   "OK"
         Height          =   375
         Left            =   7440
         TabIndex        =   90
         Top             =   1800
         Width           =   1095
      End
      Begin VB.OptionButton optSpain 
         BackColor       =   &H00000000&
         Caption         =   "SPA"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   8040
         TabIndex        =   88
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optCAN 
         BackColor       =   &H00000000&
         Caption         =   "CAN"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6600
         TabIndex        =   86
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optUSA 
         BackColor       =   &H00000000&
         Caption         =   "USA"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7320
         TabIndex        =   87
         Top             =   240
         Width           =   735
      End
      Begin VB.ComboBox cmbfrs 
         Height          =   315
         Left            =   240
         TabIndex        =   77
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox txtPrixNet 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4920
         TabIndex        =   83
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtPrixList 
         BackColor       =   &H00FFFFFF&
         DataField       =   "PRIX_LIST"
         DataSource      =   "DatCat1"
         Height          =   285
         Left            =   4920
         TabIndex        =   79
         Top             =   480
         Width           =   1095
      End
      Begin MSMask.MaskEdBox mskEscompte 
         Height          =   255
         Left            =   4920
         TabIndex        =   81
         Top             =   840
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Prix Spécial :"
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
         Left            =   3720
         TabIndex        =   84
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Image imgEU 
         Height          =   1065
         Left            =   6720
         Picture         =   "frmProjSoumMec.frx":206CE
         Stretch         =   -1  'True
         Top             =   600
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Image imgCanada 
         Height          =   1065
         Left            =   6720
         Picture         =   "frmProjSoumMec.frx":6D440
         Stretch         =   -1  'True
         Top             =   600
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Image imgSpain 
         Height          =   1065
         Left            =   6720
         Picture         =   "frmProjSoumMec.frx":C3422
         Stretch         =   -1  'True
         Top             =   600
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Prix Net :"
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
         Index           =   20
         Left            =   3720
         TabIndex        =   82
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Escompte :"
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
         Index           =   19
         Left            =   3720
         TabIndex        =   80
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Prix Listé :"
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
         Index           =   16
         Left            =   3720
         TabIndex        =   78
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Distributeur :"
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
         Index           =   14
         Left            =   240
         TabIndex        =   76
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox txtNoSoumission 
      Height          =   288
      Left            =   9000
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.ComboBox cmbChoix 
      Height          =   315
      ItemData        =   "frmProjSoumMec.frx":C58B1
      Left            =   3240
      List            =   "frmProjSoumMec.frx":C58BB
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox cmbProjSoum 
      Height          =   315
      Left            =   6000
      TabIndex        =   2
      Text            =   "cmbProjSoum"
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtNoProjSoum 
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtChoix 
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprimer 
      Caption         =   "Imprimer"
      Height          =   375
      Left            =   120
      TabIndex        =   99
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdTexte 
      Caption         =   "Texte"
      Height          =   375
      Left            =   120
      TabIndex        =   100
      Top             =   7320
      Width           =   975
   End
   Begin VB.ComboBox cmbClient 
      Height          =   315
      Left            =   6000
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   600
      Width           =   3375
   End
   Begin VB.TextBox txtClient 
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   600
      Width           =   3375
   End
   Begin VB.ComboBox cmbContact 
      Height          =   315
      Left            =   6000
      Sorted          =   -1  'True
      TabIndex        =   33
      Top             =   960
      Width           =   3375
   End
   Begin VB.TextBox txtContact 
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   960
      Width           =   3375
   End
   Begin VB.TextBox txtDescription 
      Height          =   525
      Left            =   6000
      MultiLine       =   -1  'True
      TabIndex        =   44
      Top             =   1320
      Width           =   3375
   End
   Begin VB.TextBox txtCheminPhotos 
      Height          =   285
      Left            =   6000
      TabIndex        =   48
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox txtPrixReception 
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   59
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   255
      Left            =   8400
      TabIndex        =   50
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton cmdPhotos 
      Caption         =   "Afficher"
      Height          =   255
      Left            =   8760
      TabIndex        =   51
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmdTemps 
      Caption         =   "Temps"
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   840
      Width           =   1215
   End
   Begin VB.Frame fraManuel 
      BackColor       =   &H00000000&
      Caption         =   "Manuels"
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
      Height          =   615
      Left            =   120
      TabIndex        =   38
      Top             =   1320
      Width           =   1935
      Begin VB.TextBox txtPrixManuel 
         Height          =   288
         Left            =   1320
         TabIndex        =   41
         Text            =   "0"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtNbreManuel 
         Height          =   288
         Left            =   480
         MaxLength       =   4
         TabIndex        =   39
         Text            =   "0"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Prix"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   42
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Nbre"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.ComboBox cmbSections 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmProjSoumMec.frx":C58D3
      Left            =   1080
      List            =   "frmProjSoumMec.frx":C58D5
      Style           =   2  'Dropdown List
      TabIndex        =   67
      Top             =   2640
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdAjouterSection 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3360
      TabIndex        =   68
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox cmbPieces 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmProjSoumMec.frx":C58D7
      Left            =   4800
      List            =   "frmProjSoumMec.frx":C58D9
      TabIndex        =   70
      Top             =   2640
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.ComboBox cmbTri 
      Height          =   315
      ItemData        =   "frmProjSoumMec.frx":C58DB
      Left            =   8880
      List            =   "frmProjSoumMec.frx":C58EE
      Style           =   2  'Dropdown List
      TabIndex        =   71
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton cmdTri 
      Caption         =   "Trier"
      Height          =   315
      Left            =   10800
      TabIndex        =   72
      Top             =   2700
      Width           =   975
   End
   Begin VB.CommandButton cmdRafraichir 
      Caption         =   "Rafraichir"
      Height          =   315
      Left            =   10800
      TabIndex        =   64
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox txtPrixTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   """$""# ##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3084
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   288
      Left            =   13560
      Locked          =   -1  'True
      TabIndex        =   53
      Text            =   "0"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtCommission 
      Alignment       =   1  'Right Justify
      Height          =   288
      Left            =   13560
      Locked          =   -1  'True
      TabIndex        =   46
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtImprevus 
      Alignment       =   1  'Right Justify
      Height          =   288
      Left            =   13560
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtProfit 
      Alignment       =   1  'Right Justify
      Height          =   288
      Left            =   13560
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtTotalPieces 
      Alignment       =   1  'Right Justify
      Height          =   288
      Left            =   13560
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtTotalTemps 
      Alignment       =   1  'Right Justify
      Height          =   288
      Left            =   13560
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdHistorique 
      Caption         =   "Historique des modifications"
      Height          =   495
      Left            =   120
      TabIndex        =   54
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox txtDateFacturation 
      Height          =   288
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   57
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdDateFacturation 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   58
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdLegende 
      Caption         =   "Légende"
      Height          =   375
      Left            =   1680
      TabIndex        =   55
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdBavards 
      Caption         =   "Bavard"
      Height          =   375
      Left            =   2640
      TabIndex        =   56
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdEffacerForfait 
      Caption         =   "Effacer"
      Height          =   285
      Left            =   2580
      TabIndex        =   35
      ToolTipText     =   "Efface le forfait"
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton cmdForfait 
      Caption         =   "..."
      Height          =   285
      Left            =   2100
      TabIndex        =   34
      ToolTipText     =   "Ajoute un forfait"
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox txtForfait 
      Height          =   285
      Left            =   2100
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   840
      Width           =   975
   End
   Begin MSComctlLib.ListView lvwPieces 
      Height          =   1935
      Left            =   120
      TabIndex        =   74
      Top             =   3000
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   3413
      SortKey         =   1
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "PIECE_GRB"
         Object.Width           =   2408
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "No. d'item"
         Object.Width           =   3254
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Manufacturier"
         Object.Width           =   2037
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Description française"
         Object.Width           =   7144
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Description anglaise"
         Object.Width           =   7144
      EndProperty
   End
   Begin MSComctlLib.ListView lvwSoumission 
      Height          =   4215
      Left            =   120
      TabIndex        =   73
      Top             =   3000
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   7435
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Exporter"
      Height          =   375
      Left            =   1680
      TabIndex        =   65
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton cmdReception 
      Caption         =   "Réception"
      Height          =   375
      Left            =   4320
      TabIndex        =   104
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdRechercherClient 
      Caption         =   "..."
      Height          =   315
      Left            =   9480
      TabIndex        =   121
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Forfait :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1500
      TabIndex        =   27
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblForfaitInitiale 
      BackStyle       =   0  'Transparent
      Caption         =   "Par : "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3180
      TabIndex        =   28
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblPrixSoumission 
      BackStyle       =   0  'Transparent
      Caption         =   "$ Soumission : "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5880
      TabIndex        =   62
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lblPrixReception 
      BackStyle       =   0  'Transparent
      Caption         =   "$ Réception : "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   60
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Photos : "
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5400
      TabIndex        =   49
      Top             =   1920
      Width           =   630
   End
   Begin VB.Label lblTri 
      BackStyle       =   0  'Transparent
      Caption         =   "Trier par :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8880
      TabIndex        =   63
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label lblTotalTemps 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Temps"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   12600
      TabIndex        =   6
      Top             =   120
      Width           =   885
   End
   Begin VB.Label lblImprevus 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imprévus"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   12600
      TabIndex        =   36
      Top             =   1200
      Width           =   885
   End
   Begin VB.Label lblTotalPieces 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Pièces"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   12600
      TabIndex        =   8
      Top             =   480
      Width           =   885
   End
   Begin VB.Label lblProfit 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Profit"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   12600
      TabIndex        =   29
      Top             =   840
      Width           =   885
   End
   Begin VB.Label lblPrixTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prix Total"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   12600
      TabIndex        =   52
      Top             =   1920
      Width           =   885
   End
   Begin VB.Label lblCommission 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Administration"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   12480
      TabIndex        =   45
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblNoSoumission 
      BackStyle       =   0  'Transparent
      Caption         =   "Soumission"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8040
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Contact"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5400
      TabIndex        =   32
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Client"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   11
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblSections 
      BackStyle       =   0  'Transparent
      Caption         =   "Sections :"
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
      TabIndex        =   66
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblPiece 
      BackStyle       =   0  'Transparent
      Caption         =   "Catégorie :"
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
      Left            =   3840
      TabIndex        =   69
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblProjet 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5160
      TabIndex        =   43
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblDateFacturation 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Date Facturation"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   47
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Menu mnuRightClick 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuFacturer 
         Caption         =   "Facturer"
      End
      Begin VB.Menu mnuNC 
         Caption         =   "NC"
      End
      Begin VB.Menu mnuDateRequise 
         Caption         =   "Modifier la date requise"
      End
      Begin VB.Menu mnuCommentaire 
         Caption         =   "Ajouter / Modifier le commentaire"
      End
      Begin VB.Menu mnuMauvaisPrix 
         Caption         =   "Mauvais prix"
      End
      Begin VB.Menu mnuInutile 
         Caption         =   "Matériel inutile"
      End
      Begin VB.Menu mnuTexte 
         Caption         =   "Modifier le texte"
      End
      Begin VB.Menu mnuChangerSS 
         Caption         =   "Modifier la sous-section"
      End
      Begin VB.Menu mnuFournisseur 
         Caption         =   "Modifier le fournisseur"
      End
      Begin VB.Menu mnuAnnulerCommande 
         Caption         =   "Annuler la commande"
      End
      Begin VB.Menu mnuSupprimer 
         Caption         =   "Supprimer"
      End
      Begin VB.Menu mnuAjouterPrix 
         Caption         =   "Ajouter le prix"
      End
      Begin VB.Menu mnuSortieMagasin 
         Caption         =   "Sorti du magasin"
      End
      Begin VB.Menu mnuQuantite 
         Caption         =   "Changer la quantité"
      End
   End
End
Attribute VB_Name = "FrmProjSoumMec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Index des colonnes de lvwSoumission
Private Const I_COL_SOUM_QUANTITE         As Integer = 0
Private Const I_COL_SOUM_PIECE            As Integer = 1
Private Const I_COL_SOUM_DESCR            As Integer = 2
Private Const I_COL_SOUM_MANUFACT         As Integer = 3
Private Const I_COL_SOUM_PRIX_LIST        As Integer = 4
Private Const I_COL_SOUM_ESCOMPTE         As Integer = 5
Private Const I_COL_SOUM_PRIX_NET         As Integer = 6
Private Const I_COL_SOUM_DISTRIB          As Integer = 7
Private Const I_COL_SOUM_TOTAL            As Integer = 8
Private Const I_COL_SOUM_PROFIT           As Integer = 9
Private Const I_COL_SOUM_COMMENTAIRE      As Integer = 10
Private Const I_COL_SOUM_FACTURATION      As Integer = 11
Private Const I_COL_SOUM_DATE_COMMANDE    As Integer = 12
Private Const I_COL_SOUM_DATE_REQUISE     As Integer = 13
Private Const I_COL_SOUM_NOM_COMMANDE     As Integer = 14
Private Const I_COL_SOUM_NO_SEQUENTIEL    As Integer = 15
Private Const I_COL_SOUM_PROVENANCE       As Integer = 16

Private Const I_COL_SOUMISSION_PROV       As Integer = 11

'Index des colonnes de lvwSoumission si les colonnes contenant
'des prix ne sont pas là. (SP est pour Sans Prix)
Private Const I_COL_SOUM_SP_QUANTITE      As Integer = 0
Private Const I_COL_SOUM_SP_PIECE         As Integer = 1
Private Const I_COL_SOUM_SP_DESCR         As Integer = 2
Private Const I_COL_SOUM_SP_MANUFACT      As Integer = 3
Private Const I_COL_SOUM_SP_DISTRIB       As Integer = 4
Private Const I_COL_SOUM_SP_COMMENTAIRE   As Integer = 5
Private Const I_COL_SOUM_SP_DATE_COMMANDE As Integer = 6
Private Const I_COL_SOUM_SP_DATE_REQUISE  As Integer = 7
Private Const I_COL_SOUM_SP_NOM_COMMANDE  As Integer = 8
Private Const I_COL_SOUM_SP_NO_SEQUENTIEL As Integer = 9
Private Const I_COL_SOUM_SP_PROVENANCE    As Integer = 10

Private Const I_COL_SOUMISSION_SP_PROV    As Integer = 6

'Index des colonnes de lvwPieces
Private Const I_COL_PIECES_PIECE_GRB      As Integer = 0
Private Const I_COL_PIECES_NO_ITEM        As Integer = 1
Private Const I_COL_PIECES_MANUFACT       As Integer = 2
Private Const I_COL_PIECES_DESCR_FR       As Integer = 3
Private Const I_COL_PIECES_DESCR_EN       As Integer = 4

'Index des colonnes de lvwPieceTrouve
Private Const I_COL_RECH_PIECE_GRB        As Integer = 0
Private Const I_COL_RECH_NO_ITEM          As Integer = 1
Private Const I_COL_RECH_CATEGORIE        As Integer = 2
Private Const I_COL_RECH_MANUFACT         As Integer = 3
Private Const I_COL_RECH_DESCR_FR         As Integer = 4
Private Const I_COL_RECH_DESCR_EN         As Integer = 5

'Index des colonnes de lvwFournisseur
Private Const I_COL_FRS_FRS               As Integer = 0
Private Const I_COL_FRS_PERS_RESS         As Integer = 1
Private Const I_COL_FRS_DATE              As Integer = 2
Private Const I_COL_FRS_ENTRER_PAR        As Integer = 3
Private Const I_COL_FRS_VALIDE            As Integer = 4
Private Const I_COL_FRS_PRIX_LIST         As Integer = 5
Private Const I_COL_FRS_ESCOMPTE          As Integer = 6
Private Const I_COL_FRS_PRIX_NET          As Integer = 7
Private Const I_COL_FRS_PRIX_SP           As Integer = 8
Private Const I_COL_FRS_QUOTER            As Integer = 9
Private Const I_COL_FRS_STOCK             As Integer = 10

'Index des colonnes de lvwModification
Private Const I_COL_MODIF_EMPLOYE         As Integer = 0
Private Const I_COL_MODIF_DATE            As Integer = 1
Private Const I_COL_MODIF_HEURE           As Integer = 2
Private Const I_COL_MODIF_MONTANT         As Integer = 3

'Index des colonnes de lvwBavard
Private Const I_COL_SUPP_EMPLOYE          As Integer = 0
Private Const I_COL_SUPP_DATE             As Integer = 1
Private Const I_COL_SUPP_HEURE            As Integer = 2
Private Const I_COL_SUPP_QTE              As Integer = 3
Private Const I_COL_SUPP_NO_ITEM          As Integer = 4

'Index de cmbChoix
Private Const I_IDX_SOUMISSION            As Integer = 0
Private Const I_IDX_PROJET                As Integer = 1

'Index de cmbOuvertFerme
Private Const I_CMB_OUVERT                As Integer = 0
Private Const I_CMB_TOUS                  As Integer = 1

'Constante s'il n'y a pas de sous-sections
Private Const S_PAS_SOUS_SECTION          As String = "PAS DE SOUS-SECTION"

'Valeur servant au resize du lvwSoumission si le form est agrandi
Private Const I_TOP_AFFICHAGE             As Integer = 3000
Private Const I_HEIGHT_AFFICHAGE          As Integer = 3930

'Index de cmbTri
Private Const I_CMB_PIECE_GRB             As Integer = 0
Private Const I_CMB_PIECE                 As Integer = 1
Private Const I_CMB_FABRICANT             As Integer = 2
Private Const I_CMB_DESCR_FR              As Integer = 3
Private Const I_CMB_DESCR_EN              As Integer = 4

'Énumeration servant à savoir si c'est l'affichage des soumissions ou des projets
Private Enum enumType
  TYPE_PROJET = 0
  TYPE_SOUMISSION = 1
End Enum

'Énumération servant à savoir si le form est en mode modif/ajout ou en mode
'inactif (affichage seulement)
Private Enum enumMode
  MODE_AJOUT_MODIF = 0
  MODE_INACTIF = 1
End Enum

'Énumération servant à savoir si le form est en anglais ou en francais
Private Enum enumLangage
  FRANCAIS = 0
  ANGLAIS = 1
End Enum

Private Type tyCopiePiece
  bChecked     As Boolean
  sQuantite    As String
  sPiece       As String
  sDescr       As String
  sManufact    As String
  sPrixList    As String
  sEscompte    As String
  sPrixNet     As String
  sFRS         As String
  sTotal       As String
  sProfit      As String
  sDescrTag    As String
  sPrixListTag As String
  sFRSTag      As String
  lColor       As Long
  iNoLigne     As Integer
End Type

'Variables pour le temps
Public m_bTempsProjLock          As Boolean

Public m_sTempsDessin            As String
Public m_sTempsCoupe             As String
Public m_sTempsMachinage         As String
Public m_sTempsSoudure           As String
Public m_sTempsAssemblage        As String
Public m_sTempsPeinture          As String
Public m_sTempsTest              As String
Public m_sTempsInstallation      As String
Public m_sTempsFormation         As String
Public m_sTempsGestion           As String
Public m_sTempsShipping          As String

Public m_sTempsDessinProj        As String
Public m_sTempsCoupeProj         As String
Public m_sTempsMachinageProj     As String
Public m_sTempsSoudureProj       As String
Public m_sTempsAssemblageProj    As String
Public m_sTempsPeintureProj      As String
Public m_sTempsTestProj          As String
Public m_sTempsInstallationProj  As String
Public m_sTempsFormationProj     As String
Public m_sTempsGestionProj       As String
Public m_sTempsShippingProj      As String
Public m_sTempsPrototypeProj      As String

Public m_sTempsDessinConc        As String
Public m_sTempsCoupeConc         As String
Public m_sTempsMachinageConc     As String
Public m_sTempsSoudureConc       As String
Public m_sTempsAssemblageConc    As String
Public m_sTempsPeintureConc      As String
Public m_sTempsTestConc          As String
Public m_sTempsInstallationConc  As String
Public m_sTempsFormationConc     As String
Public m_sTempsGestionConc       As String
Public m_sTempsShippingConc      As String
Public m_sTempsPrototypeConc      As String

Public m_sNbrePersonne           As String
Public m_sTempsHebergement       As String
Public m_sTempsRepas             As String
Public m_sTempsTransport         As String
Public m_sTempsUniteMobile       As String
Public m_sPrixEmballage          As String

Public m_sTauxHebergement1       As String
Public m_sTauxHebergement2       As String
Public m_sTauxRepas              As String
Public m_sTauxTransport          As String
Public m_sTauxUniteMobile        As String

Public m_sTauxDessin             As String
Public m_sTauxCoupe              As String
Public m_sTauxMachinage          As String
Public m_sTauxSoudure            As String
Public m_sTauxAssemblage         As String
Public m_sTauxPeinture           As String
Public m_sTauxTest               As String
Public m_sTauxInstallation       As String
Public m_sTauxFormation          As String
Public m_sTauxGestion            As String
Public m_sTauxShipping           As String

'Pour savoir si l'écran des temps a déjà été ouvert
Public m_bTempsDejaOuvert        As Boolean

'Pour savoir si le form a déjà été sur l'événement resize
Private m_bResize                As Boolean

'Variables pour la configurations
Private m_sProfit                As String
Private m_sCommission            As String
Private m_sImprevue              As String

'Modes du form
Private m_bModeAjout             As Boolean
Private m_bModeAffichage         As Boolean

'Pour avoir une sous-section par défaut
Private m_sSousSection           As String

'Pour le tri de lvwPieces
Private m_sTri                   As String

'Pour savoir quelle colonne trier
Private m_iCol                   As Integer

'Pour savoir si le form affiche les projets ou les soumissions
Private m_eType                  As enumType
Private m_eMode                  As enumMode

'Pour savoir si les prix sont cachés ou non
Public m_bDroitPrix              As Boolean

'Pour ne pas être obligé d'ouvrir le recordset à chaque fois
Private m_bModifProj             As Boolean
Private m_bModifSoum             As Boolean
Private m_bModifBonCommande      As Boolean

'Variable pour savoir si l'utilisateur a le droit de voir le combo ou non
Private m_bComboChoix            As Boolean

'Pour faire afficher le dernier enregitrement visionné après un ajout ou une
'modification
Private m_sAncienProjSoum        As String

'Pour savoir si c'est en francais ou en anglais
Private m_eLangage               As enumLangage

'Pour savoir si l'utilisateur a le droit de supprimer des projets
Private m_bSupprimer             As Boolean

Private m_bPieceInutile          As Boolean

Public m_bAnnulerChemin          As Boolean

Public m_sChemin                 As String

Private m_bRecherchePiece        As Boolean
 
Private m_bExtra                 As Boolean

'Pour ne pas resetter le langage après un enregistrement
Private m_bEnregistrement        As Boolean
 
Private m_bMauvaisPrix           As Boolean

Private m_collDateSupp           As Collection
Private m_collHeureSupp          As Collection
Private m_collQteSupp            As Collection
Private m_collNoItemSupp         As Collection

Private m_bChangementFRS         As Boolean

Private m_sTexteRecherche        As String

Private m_arr_tyCopie()          As tyCopiePiece

Private m_iNbreCopie             As Integer

Public m_bModifFournisseurBC     As Boolean

Private m_sLiaison               As String
Private m_bMonthViewHasFocus     As Boolean

Public m_bTransfertJobCancel     As Boolean

Private m_bChangementChoix       As Boolean 'Pour empêcher l'événement cmbOuvertFerme_Click quand cmbChoix change

Public m_bValide                 As Boolean 'Résultat de frmValiderSuppression
Public intdummie                 As Integer 'va servir à annuler l'impression des rapports dans la sub ImprimerProjSoum
Public bTrigger                  As Boolean
'**********************************************************************************************************
'AJOUT PAR GAÉTAN GINGRAS 07 FÉVRIER 2010
'**********************************************************************************************************
Public bFlag As Boolean 'pour garder en mémoire si on désir achiffer les dates de réception et de commande
'**********************************************************************************************************

Public Function PeutFermer() As Boolean

5       On Error GoTo AfficherErreur

10      If m_eMode = MODE_INACTIF Then
15        PeutFermer = True
20      Else
25        PeutFermer = False
30      End If

35      Exit Function

AfficherErreur:

40      woups "frmProjSoumMec", "PeutFermer", Err, Erl
End Function

Private Sub InitialiserVariables(ByVal sNoProjSoum As String)

5       On Error GoTo AfficherErreur

        'Initialisation des variables comprises dans la configuration
10      Dim rstConfig   As ADODB.Recordset
15      Dim rstProjSoum As ADODB.Recordset
  
20      Set rstConfig = New ADODB.Recordset
25      Set rstProjSoum = New ADODB.Recordset
  
30      Call rstConfig.Open("SELECT * FROM GRB_Config", g_connData, adOpenDynamic, adLockOptimistic)

35      If m_eType = TYPE_PROJET Then
40        Call rstProjSoum.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)
45      Else
50        Call rstProjSoum.Open("SELECT * FROM GRB_SoumissionMec WHERE IDSoumission = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)
55      End If

60      If Not rstProjSoum.EOF Then
65        m_sProfit = rstProjSoum.Fields("Profit")
70        m_sCommission = rstProjSoum.Fields("Commission")
75        m_sImprevue = rstProjSoum.Fields("Imprevue")
80      Else
85        m_sProfit = rstConfig.Fields("ProfitMec")
90        m_sCommission = rstConfig.Fields("Commission")
95        m_sImprevue = rstConfig.Fields("Imprévus")
100     End If
    
105     Call rstConfig.Close
110     Set rstConfig = Nothing

115     Call rstProjSoum.Close
120     Set rstProjSoum = Nothing

125     Exit Sub

AfficherErreur:

130     woups "frmProjSoumMec", "InitialiserVariables", Err, Erl
End Sub

Private Sub ActiverBoutonsGroupe()

5       On Error GoTo AfficherErreur

        'Activation des boutons d'après le groupe
10      Dim bModif As Boolean
  
        'Si l'utilisateur a le droit d'affichage sur les projets et les soumissions
15      If g_bAffichageProjetsMec = True And g_bAffichageSoumissionsMec = True Then
          'On affiche cmbChoix
20        cmbChoix.Visible = True
      
          'Cette variable sert à savoir si l'utilisateur a le droit de voir le combo
25        m_bComboChoix = True
      
          'Type d'affichage
30        m_eType = TYPE_PROJET
    
          'Champs pour la modification
35        bModif = g_bModificationProjetsMec
40      Else
          'On cache cmbChoix
45        cmbChoix.Visible = False
                   
          'Cette variable sert à savoir si l'utilisateur a le droit de voir le combo
50        m_bComboChoix = False
                  
          'Si l'utilisateur a le droit d'affichage sur les projets
55        If g_bAffichageProjetsMec = True Then
            'Le seul choix possible est Projet
60          txtChoix.Text = "Projet"
        
            'Le type d'affichage
80          m_eType = TYPE_PROJET
        
            'Champs pour la modification
85          bModif = g_bModificationProjetsMec
90        Else
            'Le seul choix possible est Soumission
95          txtChoix.Text = "Soumission"
                              
            'Type d'affichage
100         m_eType = TYPE_SOUMISSION
                              
            'Champs pour la modification
105         bModif = g_bModificationSoumissionsMec
110       End If
115     End If
         
120     m_bModifProj = g_bModificationProjetsMec
125     m_bModifSoum = g_bModificationSoumissionsMec
130     m_bModifBonCommande = g_bModificationBC
135     m_bSupprimer = g_bSuppressionProjets
    
140     Cmdajouter.Enabled = bModif
145     cmdsupprimer.Enabled = bModif
150     cmdModifier.Enabled = bModif
155     cmdCopier.Enabled = bModif
160     cmdCreerProjet.Enabled = bModif
165     cmdBonCommande.Enabled = m_bModifBonCommande
170     cmdImprimer.Enabled = bModif
175     cmdDemande.Enabled = bModif
180     cmdAnglaisFrancais.Enabled = bModif
185     cmdExtra.Enabled = bModif
190     cmdSupprimerFRS.Visible = g_bModificationCatalogueMec
195     cmdRetour.Enabled = g_bModificationRetourMarchandise
200     cmdReception.Enabled = g_bModificationReception

205     Exit Sub

AfficherErreur:

210     woups "frmProjSoumMec", "ActiverBoutonsGroupe", Err, Erl
End Sub

Private Sub AfficherProjSoum(ByVal sNoProjSoum As String)

5       On Error GoTo AfficherErreur
        
10      m_bPieceInutile = False
15      m_bChangementFRS = False
20      m_bRecherchePiece = False
        
        'Remet en mode affichage le projet ou la soumission voulue
25      m_bModeAffichage = True
    
        'Vide les champs
30      Call ViderChamps
  
        'Rempli le combo
35      Call RemplirComboProjSoum(sNoProjSoum)
  
        'Barre les champs
40      Call BarrerChamps(True)
  
45      lvwSoumission.Top = I_TOP_AFFICHAGE
50      lvwSoumission.Height = Me.Height - I_HEIGHT_AFFICHAGE

55      Exit Sub

AfficherErreur:

60      woups "frmProjSoumMec", "AfficherProjSoum", Err, Erl
End Sub

Private Sub AfficherControles(ByVal eMode As enumMode)

5       On Error GoTo AfficherErreur

        'Affichage des boutons selon si c'est un ajout/modif ou un affichage
10      Dim bAjouter         As Boolean
15      Dim bModifier        As Boolean
20      Dim bSupprimer       As Boolean
25      Dim bEnregistrer     As Boolean
30      Dim bAnnuler         As Boolean
35      Dim bFermer          As Boolean
40      Dim bImprimer        As Boolean
45      Dim bCmbClient       As Boolean
50      Dim bCmbContact      As Boolean
55      Dim bCmbProjSoum     As Boolean
60      Dim bCmbTransport    As Boolean
65      Dim bCmbChoix        As Boolean
70      Dim bCmbOuvertFerme  As Boolean
75      Dim bSection         As Boolean
80      Dim bPieces          As Boolean
85      Dim bTexte           As Boolean
90      Dim bCreerProjet     As Boolean
95      Dim bHistorique      As Boolean
100     Dim bCopier          As Boolean
105     Dim bBonCommande     As Boolean
110     Dim bTri             As Boolean
115     Dim bDemande         As Boolean
120     Dim bExtra           As Boolean
125     Dim bCatalogue       As Boolean
130     Dim bBrowseChemin    As Boolean
135     Dim bInutile         As Boolean
140     Dim bMauvaisPrix     As Boolean
145     Dim bRapportFact     As Boolean
150     Dim bDateFacture     As Boolean
155     Dim bSortiMagasin    As Boolean
160     Dim bRetour          As Boolean
165     Dim bForfait         As Boolean
170     Dim bExport          As Boolean
175     Dim bReception       As Boolean
180     Dim bAnglaisFrancais As Boolean
185     Dim bRechercheClient As Boolean
  
190     m_eMode = eMode
  
195     Select Case eMode
          Case MODE_AJOUT_MODIF:
200         bEnregistrer = True
205         bAnnuler = True
210         bSection = True
215         bPieces = True
220         bTexte = True
225         bTri = True

230         If (m_eType = TYPE_SOUMISSION) Or (m_eType = TYPE_PROJET And Mid$(txtNoProjSoum.Text, 3, 1) <> "3") Then
235           bCmbClient = True
240           bCmbContact = True
245           bRechercheClient = True
250         End If

255         bCmbTransport = True
260         bCatalogue = True
265         bBrowseChemin = True
270         bMauvaisPrix = True
275         bForfait = True

280         If m_eType = TYPE_PROJET Then
285           bInutile = True

290           If g_bModificationReception = True Then
295             bSortiMagasin = True
300           End If

305           If g_bModificationFacturation = True Then
310             bDateFacture = True
315           End If
320         End If
        
325      Case MODE_INACTIF:
330         bModifier = True
335         bFermer = True
340         bImprimer = True
345         bCmbProjSoum = True
350         bCmbChoix = True
355         bCmbOuvertFerme = True
360         bHistorique = True
365         bDemande = True
370         bExport = True
375         bAnglaisFrancais = True
380         bAjouter = True
               
385         If m_eType = TYPE_PROJET Then
390           bBonCommande = True
395           bExtra = True

400           If g_bModificationRetourMarchandise = True Then
405             bRetour = True
410           End If

415           If g_bModificationFacturation = True Then
420             bRapportFact = True
425           End If

430           If g_bModificationReception = True Then
435             bReception = True
440           End If

445           If m_bSupprimer = True Then
450             bSupprimer = True
455           End If
460         Else
465           bSupprimer = True
470           bCopier = True
      
475           If VerifierSiDejaProjet = False Then
480             bCreerProjet = True
485           End If
490         End If
495     End Select

500     Cmdajouter.Visible = bAjouter
505     cmdModifier.Visible = bModifier
510     cmdsupprimer.Visible = bSupprimer
515     cmdEnregistrer.Visible = bEnregistrer
520     cmdAnnuler.Visible = bAnnuler
525     Cmdfermer.Visible = bFermer
530     cmdImprimer.Visible = bImprimer
535     cmdRapportFACT.Visible = bRapportFact
540     cmdTexte.Visible = bTexte
545     cmdHistorique.Visible = bHistorique
550     cmdCopier.Visible = bCopier
555     cmdBonCommande.Visible = bBonCommande
560     cmdDemande.Visible = bDemande
565     cmdCreerProjet.Visible = bCreerProjet
570     cmdExtra.Visible = bExtra
575     cmdCatalogue.Visible = bCatalogue
580     cmdBrowse.Visible = bBrowseChemin
585     cmdMaterielInutile.Visible = bInutile
590     cmdMauvaisPrix.Visible = bMauvaisPrix
595     cmdSortieMagasin.Visible = bSortiMagasin
600     cmdRetour.Visible = bRetour
605     cmdForfait.Visible = bForfait
610     cmdEffacerForfait.Visible = bForfait
615     cmdExport.Visible = bExport
620     cmdReception.Visible = bReception
625     cmdAnglaisFrancais.Visible = bAnglaisFrancais

630     lblDateFacturation.Visible = bDateFacture
635     txtDateFacturation.Visible = bDateFacture
640     cmdDateFacturation.Visible = bDateFacture

645     cmbclient.Visible = bCmbClient
650     txtClient.Visible = Not bCmbClient

655     cmbContact.Visible = bCmbContact
660     txtcontact.Visible = Not bCmbContact

        'Si on a le droit d'affiche le combo
665     If m_bComboChoix = True Then
670       cmbChoix.Visible = bCmbChoix
675       txtChoix.Visible = Not bCmbChoix
680     End If

685     cmbOuvertFerme.Visible = bCmbOuvertFerme

690     cmbProjSoum.Visible = bCmbProjSoum
695     txtNoProjSoum.Visible = Not bCmbProjSoum

700     lblSections.Visible = bSection
705     cmbSections.Visible = bSection
710     cmdAjouterSection.Visible = bSection

715     lblPiece.Visible = bPieces
720     cmbPieces.Visible = bPieces
725     lvwPieces.Visible = bPieces

730     lblTri.Visible = bTri
735     cmbTri.Visible = bTri
740     cmdTri.Visible = bTri
745     cmdRafraichir.Visible = bTri

750     lblPrixTotal.Visible = m_bDroitPrix
755     lblCommission.Visible = m_bDroitPrix
760     lblProfit.Visible = m_bDroitPrix
765     lblImprevus.Visible = m_bDroitPrix
770     lblTotalPieces.Visible = m_bDroitPrix
775     lblTotalTemps.Visible = m_bDroitPrix

780     txtPrixTotal.Visible = m_bDroitPrix
785     txtCommission.Visible = m_bDroitPrix
790     txtProfit.Visible = m_bDroitPrix
795     txtImprevus.Visible = m_bDroitPrix
800     txtTotalPieces.Visible = m_bDroitPrix
805     txtTotalTemps.Visible = m_bDroitPrix

810     cmdRechercherClient.Visible = bRechercheClient

815     Exit Sub

AfficherErreur:

820     woups "frmProjSoumMec", "AfficherControles", Err, Erl
End Sub

Private Sub cmbChoix_Click()

5       On Error GoTo AfficherErreur

10      Dim bModif          As Boolean
15      Dim iCmbOuvertFerme As Integer
    
20      Screen.MousePointer = vbHourglass
      
25      txtChoix.Text = cmbChoix.Text

        'Met les CheckBoxes sur le ListView
30      lvwSoumission.Checkboxes = True

35      If cmbChoix.ListIndex = I_IDX_SOUMISSION Then
          'Change le type
40        m_eType = TYPE_SOUMISSION
    
45        m_bChangementChoix = True

50        iCmbOuvertFerme = cmbOuvertFerme.ListIndex

55        Call cmbOuvertFerme.Clear

60        Call cmbOuvertFerme.AddItem("Ouvertes")
65        Call cmbOuvertFerme.AddItem("Toutes")

70        cmbOuvertFerme.ListIndex = iCmbOuvertFerme

75        m_bChangementChoix = False
    
80        bModif = m_bModifSoum
        
          'Cache la soumission
85        lblNoSoumission.Visible = False
90        txtNoSoumission.Visible = False

          'Cache Prix Réception
95        lblPrixReception.Visible = False
100       txtPrixReception.Visible = False

          'Cache Prix Soumission
105       lblPrixSoumission.Visible = False
110       txtPrixSoumission.Visible = False
115     Else
          'Change le type
120       m_eType = TYPE_PROJET
    
125       m_bChangementChoix = True
    
130       iCmbOuvertFerme = cmbOuvertFerme.ListIndex

135       Call cmbOuvertFerme.Clear

140       Call cmbOuvertFerme.AddItem("Ouverts")
145       Call cmbOuvertFerme.AddItem("Tous")

150       cmbOuvertFerme.ListIndex = iCmbOuvertFerme
    
155       m_bChangementChoix = False
    
160       bModif = m_bModifProj
   
          'Affiche la soumission
165       lblNoSoumission.Visible = True
170       txtNoSoumission.Visible = True

          'Affiche Prix Réception
175       lblPrixReception.Visible = True
180       txtPrixReception.Visible = True

          'Affiche Prix Soumission
185       lblPrixSoumission.Visible = True
190       txtPrixSoumission.Visible = True

195       txtDateFacturation.Text = ConvertDate(Date)
200     End If
    
        'Active ou désactive les boutons de modification selon
        'le groupe auquel l'utilisateur appartient
205     cmdModifier.Enabled = bModif
210     cmdsupprimer.Enabled = bModif
215     Cmdajouter.Enabled = bModif
220     cmdCopier.Enabled = bModif
225     cmdCreerProjet.Enabled = bModif
230     cmdBonCommande.Enabled = m_bModifBonCommande
235     cmdImprimer.Enabled = bModif
240     cmdDemande.Enabled = bModif
245     cmdAnglaisFrancais.Enabled = bModif
250     cmdExtra.Enabled = bModif
  
        'Ajoute les colonnes selon le groupe
255     Call RemplirColonnes
  
260     m_bModeAffichage = True
  
        'Vide les champs
265     Call ViderChamps
  
        'Barre les champs
270     Call BarrerChamps(True)
  
        'Rempli le combo
275     Call RemplirComboProjSoum(vbNullString)
  
        'Affiche les controles pour le mode inactif
280     Call AfficherControles(MODE_INACTIF)
  
285     Call PositionnerBoutons
  
290     Screen.MousePointer = vbDefault

295     Exit Sub

AfficherErreur:

300     woups "frmProjSoumMec", "cmbChoix_Click", Err, Erl
End Sub

Private Sub cmbclient_Click()

5       On Error GoTo AfficherErreur

        'Rempli le combo des contacts selon le client choisi
10      Call RemplirComboContacts

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMec", "cmbclient_Click", Err, Erl
End Sub

Private Sub cmbOuvertFerme_Click()
        
5       On Error GoTo AfficherErreur

10      If cmbChoix.ListIndex <> -1 Then
15        If m_bChangementChoix = False Then
20          Call RemplirComboProjSoum("")
25        End If
30      End If

35      Exit Sub

AfficherErreur:

40      woups "FrmProjSoumElec", "cmbOuvertFerme_Click", Err, Erl
End Sub

Private Sub cmbPieces_Click()

5       On Error GoTo AfficherErreur

        'Rempli lvwPieces selon la catégorie de pièce choisie
10      Call RemplirListViewPieces

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMec", "cmbPieces_Click", Err, Erl
End Sub

Private Sub cmbProjSoum_Click()

5       On Error GoTo AfficherErreur

10      Dim rstProjSoum As ADODB.Recordset
15      Dim rstOuvert   As ADODB.Recordset
20      Dim iCompteur   As Integer
25      Dim sNomClient  As String
30      Dim sNomContact As String
35      Dim sNumero     As String
40      Dim bTrouve     As Boolean
  
45      Screen.MousePointer = vbHourglass
  
50      If cmbProjSoum.Text <> "" Then
55        sNumero = txtNoProjSoum.Text

60        txtNoProjSoum.Text = cmbProjSoum.Text

65        Call InitialiserVariables(txtNoProjSoum.Text)

70        If m_bEnregistrement = False Then
75          m_eLangage = FRANCAIS

80          cmdAnglaisFrancais.Caption = "Anglais"
85        End If

90        Set rstProjSoum = New ADODB.Recordset

95        If m_eType = TYPE_PROJET Then
100         Call rstProjSoum.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
105       Else
110         Call rstProjSoum.Open("SELECT * FROM GRB_SoumissionMec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
115       End If

120       If rstProjSoum.Fields("Modification") = True And rstProjSoum.Fields("Par") = g_sEmploye Then
125         cmdReset.Visible = True
130       End If

135       Call InitialiserTempsTaux(False)
  
140       If m_eType = TYPE_SOUMISSION Then
            'Si la soumission n'est pas assigné à un projet
145         If VerifierSiDejaProjet = False Then
              'On affiche le bouton cmdCreerProjet
150           cmdCreerProjet.Visible = True
155         Else
160           cmdCreerProjet.Visible = False
165         End If
170       End If
    
          'Rempli les valeurs de la soumission ou du projet sélectionné
175       Call RemplirProjSoum

          'Le temps calculé dans le projet est le temps réel, c'est pourquoi il faut le recalculer
          'puisque le temps varie souvent
180       If m_eType = TYPE_PROJET Then
185         Set rstOuvert = New ADODB.Recordset

190         Call rstOuvert.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

195         If rstOuvert.Fields("Ouvert") = True Then
200           m_bModeAffichage = False

205           Call CalculerPrix

210           m_bModeAffichage = True

215           rstProjSoum.Fields("total_Commission") = txtCommission.Text
220           rstProjSoum.Fields("Total_Profit") = txtProfit.Text
225           rstProjSoum.Fields("PrixTotal") = txtPrixTotal.Text
230           rstProjSoum.Fields("Total_piece") = txtTotalPieces.Text
235           rstProjSoum.Fields("total_imprevue") = txtImprevus.Text
240           rstProjSoum.Fields("Total_Temps") = txtTotalTemps.Text

245           Call rstProjSoum.Update
250         End If
255       End If
  
260       Call rstProjSoum.Close
  
265       sNomClient = txtClient.Text
270       sNomContact = txtcontact.Text
    
          'Pour choisir le bon client dans le combo des clients
275       For iCompteur = 0 To cmbclient.ListCount - 1
280         If cmbclient.LIST(iCompteur) = sNomClient Then
285           cmbclient.ListIndex = iCompteur
        
290           bTrouve = True
        
295           Exit For
300         End If
305       Next

310       If bTrouve = False Then
315         Call RemplirComboClients(vbNullString)

320         For iCompteur = 0 To cmbclient.ListCount - 1
325           If cmbclient.LIST(iCompteur) = sNomClient Then
330             cmbclient.ListIndex = iCompteur

335             Exit For
340           End If
345         Next
350       End If
      
          'Pour choisir le bon contact dans le combo des contacts
355       For iCompteur = 0 To cmbContact.ListCount - 1
360         If cmbContact.LIST(iCompteur) = sNomContact Then
365           cmbContact.ListIndex = iCompteur
        
370           Exit For
375         End If
380       Next
385     End If

390     Call CalculerPrixReception

395     If m_eType = TYPE_PROJET Then
400       Call rstProjSoum.Open("SELECT PrixRéception FROM GRB_ProjetMec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
    
405       rstProjSoum.Fields("PrixRéception") = txtPrixReception.Text
 
410       Call rstProjSoum.Update

415       Call rstProjSoum.Close
420     End If
    
425     Set rstProjSoum = Nothing
    
430     Screen.MousePointer = vbDefault

435     Exit Sub

AfficherErreur:

440     woups "frmProjSoumMec", "cmbProjSoum_Click", Err, Erl, txtNoProjSoum.Text)
End Sub

Private Sub InitialiserTempsTaux(ByVal bEmpty As Boolean)

5       On Error GoTo AfficherErreur

        'Pour initialiser les temps et les taux horaires
10      Dim rstProjSoum As ADODB.Recordset
15      Dim sTable      As String
20      Dim sChamps     As String
  
25      m_bTempsDejaOuvert = False
    
30      If bEmpty = True Then
35        m_sTempsDessin = "0"
40        m_sTempsCoupe = "0"
45        m_sTempsMachinage = "0"
50        m_sTempsSoudure = "0"
55        m_sTempsAssemblage = "0"
60        m_sTempsPeinture = "0"
65        m_sTempsTest = "0"
70        m_sTempsInstallation = "0"
75        m_sTempsFormation = "0"
80        m_sTempsGestion = "0"
85        m_sTempsShipping = "0"

90        m_sNbrePersonne = "0"
95        m_sTempsHebergement = "0"
100       m_sTempsRepas = "0"
105       m_sTempsTransport = "0"
110       m_sTempsUniteMobile = "0"
115       m_sPrixEmballage = "0"

120       m_sTauxDessin = "0"
125       m_sTauxCoupe = "0"
130       m_sTauxMachinage = "0"
135       m_sTauxSoudure = "0"
140       m_sTauxAssemblage = "0"
145       m_sTauxPeinture = "0"
150       m_sTauxTest = "0"
155       m_sTauxInstallation = "0"
160       m_sTauxFormation = "0"
165       m_sTauxGestion = "0"
170       m_sTauxShipping = "0"

175       m_sTauxHebergement1 = "0"
180       m_sTauxHebergement2 = "0"
185       m_sTauxRepas = "0"
190       m_sTauxTransport = "0"
195       m_sTauxUniteMobile = "0"

200       m_bTempsProjLock = False
205     Else
210       If m_eType = TYPE_PROJET Then
215         sTable = "GRB_ProjetMec"
220         sChamps = "IDProjet"
225       Else
230         sTable = "GRB_SoumissionMec"
235         sChamps = "IDSoumission"
240       End If

245       Set rstProjSoum = New ADODB.Recordset
  
250       Call rstProjSoum.Open("SELECT * FROM " & sTable & " WHERE " & sChamps & " = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
      
255       If m_eType = TYPE_SOUMISSION Then
260         If Not IsNull(rstProjSoum.Fields("TempsDessin")) Then
265           m_sTempsDessin = rstProjSoum.Fields("TempsDessin")
270         Else
275           m_sTempsDessin = "0"
280         End If

285         If Not IsNull(rstProjSoum.Fields("TempsCoupe")) Then
290           m_sTempsCoupe = rstProjSoum.Fields("TempsCoupe")
295         Else
300           m_sTempsCoupe = "0"
305         End If

310         If Not IsNull(rstProjSoum.Fields("TempsMachinage")) Then
315           m_sTempsMachinage = rstProjSoum.Fields("TempsMachinage")
320         Else
325           m_sTempsMachinage = "0"
330         End If

335         If Not IsNull(rstProjSoum.Fields("TempsSoudure")) Then
340           m_sTempsSoudure = rstProjSoum.Fields("TempsSoudure")
345         Else
350           m_sTempsSoudure = "0"
355         End If

360         If Not IsNull(rstProjSoum.Fields("TempsAssemblage")) Then
365           m_sTempsAssemblage = rstProjSoum.Fields("TempsAssemblage")
370         Else
375           m_sTempsAssemblage = "0"
380         End If

385         If Not IsNull(rstProjSoum.Fields("TempsPeinture")) Then
390           m_sTempsPeinture = rstProjSoum.Fields("TempsPeinture")
395         Else
400           m_sTempsPeinture = "0"
405         End If

410         If Not IsNull(rstProjSoum.Fields("TempsTest")) Then
415           m_sTempsTest = rstProjSoum.Fields("TempsTest")
420         Else
425           m_sTempsTest = "0"
430         End If

435         If Not IsNull(rstProjSoum.Fields("TempsInstallation")) Then
440           m_sTempsInstallation = rstProjSoum.Fields("TempsInstallation")
445         Else
450           m_sTempsInstallation = "0"
455         End If

460         If Not IsNull(rstProjSoum.Fields("TempsFormation")) Then
465           m_sTempsFormation = rstProjSoum.Fields("TempsFormation")
470         Else
475           m_sTempsFormation = "0"
480         End If

485         If Not IsNull(rstProjSoum.Fields("TempsGestion")) Then
490           m_sTempsGestion = rstProjSoum.Fields("TempsGestion")
495         Else
500           m_sTempsGestion = "0"
505         End If

510         If Not IsNull(rstProjSoum.Fields("TempsShipping")) Then
515           m_sTempsShipping = rstProjSoum.Fields("TempsShipping")
520         Else
525           m_sTempsShipping = "0"
530         End If
535       Else
540         Call InitialiserTempsReel

545         m_bTempsProjLock = rstProjSoum.Fields("TempsProjBarré")
                
550         If Not IsNull(rstProjSoum.Fields("TempsDessinProj")) Then
555           m_sTempsDessinProj = rstProjSoum.Fields("TempsDessinProj")
560         Else
565           m_sTempsDessinProj = "0"
570         End If

575         If Not IsNull(rstProjSoum.Fields("TempsCoupeProj")) Then
580           m_sTempsCoupeProj = rstProjSoum.Fields("TempsCoupeProj")
585         Else
590           m_sTempsCoupeProj = "0"
595         End If

600         If Not IsNull(rstProjSoum.Fields("TempsMachinageProj")) Then
605           m_sTempsMachinageProj = rstProjSoum.Fields("TempsMachinageProj")
610         Else
615           m_sTempsMachinageProj = "0"
620         End If

625         If Not IsNull(rstProjSoum.Fields("TempsSoudureProj")) Then
630           m_sTempsSoudureProj = rstProjSoum.Fields("TempsSoudureProj")
635         Else
640           m_sTempsSoudureProj = "0"
645         End If

650         If Not IsNull(rstProjSoum.Fields("TempsAssemblageProj")) Then
655           m_sTempsAssemblageProj = rstProjSoum.Fields("TempsAssemblageProj")
660         Else
665           m_sTempsAssemblageProj = "0"
670         End If

675         If Not IsNull(rstProjSoum.Fields("TempsPeintureProj")) Then
680           m_sTempsPeintureProj = rstProjSoum.Fields("TempsPeintureProj")
685         Else
690           m_sTempsPeintureProj = "0"
695         End If

700         If Not IsNull(rstProjSoum.Fields("TempsTestProj")) Then
705           m_sTempsTestProj = rstProjSoum.Fields("TempsTestProj")
710         Else
715           m_sTempsTestProj = "0"
720         End If

725         If Not IsNull(rstProjSoum.Fields("TempsInstallationProj")) Then
730           m_sTempsInstallationProj = rstProjSoum.Fields("TempsInstallationProj")
735         Else
740           m_sTempsInstallationProj = "0"
745         End If

750         If Not IsNull(rstProjSoum.Fields("TempsFormationProj")) Then
755           m_sTempsFormationProj = rstProjSoum.Fields("TempsFormationProj")
760         Else
765           m_sTempsFormationProj = "0"
770         End If

775         If Not IsNull(rstProjSoum.Fields("TempsGestionProj")) Then
780           m_sTempsGestionProj = rstProjSoum.Fields("TempsGestionProj")
785         Else
790           m_sTempsGestionProj = "0"
795         End If

800         If Not IsNull(rstProjSoum.Fields("TempsShippingProj")) Then
805           m_sTempsShippingProj = rstProjSoum.Fields("TempsShippingProj")
810         Else
815           m_sTempsShippingProj = "0"
820         End If

825         If Not IsNull(rstProjSoum.Fields("TempsDessinConc")) Then
830           m_sTempsDessinConc = rstProjSoum.Fields("TempsDessinConc")
835         Else
840           m_sTempsDessinConc = "0"
845         End If

850         If Not IsNull(rstProjSoum.Fields("TempsCoupeConc")) Then
855           m_sTempsCoupeConc = rstProjSoum.Fields("TempsCoupeConc")
860         Else
865           m_sTempsCoupeConc = "0"
870         End If

875         If Not IsNull(rstProjSoum.Fields("TempsMachinageConc")) Then
880           m_sTempsMachinageConc = rstProjSoum.Fields("TempsMachinageConc")
885         Else
890           m_sTempsMachinageConc = "0"
895         End If

900         If Not IsNull(rstProjSoum.Fields("TempsSoudureConc")) Then
905           m_sTempsSoudureConc = rstProjSoum.Fields("TempsSoudureConc")
910         Else
915           m_sTempsSoudureConc = "0"
920         End If

925         If Not IsNull(rstProjSoum.Fields("TempsAssemblageConc")) Then
930           m_sTempsAssemblageConc = rstProjSoum.Fields("TempsAssemblageConc")
935         Else
940           m_sTempsAssemblageConc = "0"
945         End If

950         If Not IsNull(rstProjSoum.Fields("TempsPeintureConc")) Then
955           m_sTempsPeintureConc = rstProjSoum.Fields("TempsPeintureConc")
960         Else
965           m_sTempsPeintureConc = "0"
970         End If

975         If Not IsNull(rstProjSoum.Fields("TempsTestConc")) Then
980           m_sTempsTestConc = rstProjSoum.Fields("TempsTestConc")
985         Else
990           m_sTempsTestConc = "0"
995         End If

1000        If Not IsNull(rstProjSoum.Fields("TempsInstallationConc")) Then
1005          m_sTempsInstallationConc = rstProjSoum.Fields("TempsInstallationConc")
1010        Else
1015          m_sTempsInstallationConc = "0"
1020        End If

1025        If Not IsNull(rstProjSoum.Fields("TempsFormationConc")) Then
1030          m_sTempsFormationConc = rstProjSoum.Fields("TempsFormationConc")
1035        Else
1040          m_sTempsFormationConc = "0"
1045        End If

1050        If Not IsNull(rstProjSoum.Fields("TempsGestionConc")) Then
1055          m_sTempsGestionConc = rstProjSoum.Fields("TempsGestionConc")
1060        Else
1065          m_sTempsGestionConc = "0"
1070        End If

1075        If Not IsNull(rstProjSoum.Fields("TempsShippingConc")) Then
1080          m_sTempsShippingConc = rstProjSoum.Fields("TempsShippingConc")
1085        Else
1090          m_sTempsShippingConc = "0"
1095        End If
1100      End If

1105      If m_eType = TYPE_PROJET Then
1110        m_sNbrePersonne = "0"
1115        m_sTempsHebergement = "0"
1120        m_sTempsRepas = "0"
1125        m_sTempsTransport = "0"
1130        m_sTempsUniteMobile = "0"
1135      Else
1140        If Not IsNull(rstProjSoum.Fields("NbrePersonne")) Then
1145          m_sNbrePersonne = rstProjSoum.Fields("NbrePersonne")
1150        Else
1155          m_sNbrePersonne = "0"
1160        End If
            
1165        If Not IsNull(rstProjSoum.Fields("TempsHebergement")) Then
1170          m_sTempsHebergement = rstProjSoum.Fields("TempsHebergement")
1175        Else
1180          m_sTempsHebergement = "0"
1185        End If

1190        If Not IsNull(rstProjSoum.Fields("TempsRepas")) Then
1195          m_sTempsRepas = rstProjSoum.Fields("TempsRepas")
1200        Else
1205          m_sTempsRepas = "0"
1210        End If

1215        If Not IsNull(rstProjSoum.Fields("TempsTransport")) Then
1220          m_sTempsTransport = rstProjSoum.Fields("TempsTransport")
1225        Else
1230          m_sTempsTransport = "0"
1235        End If

1240        If Not IsNull(rstProjSoum.Fields("TempsUniteMobile")) Then
1245          m_sTempsUniteMobile = rstProjSoum.Fields("TempsUniteMobile")
1250        Else
1255          m_sTempsUniteMobile = "0"
1260        End If
1265      End If

1270      m_sPrixEmballage = rstProjSoum.Fields("PrixEmballage")

1275      If Not IsNull(rstProjSoum.Fields("TauxDessin")) Then
1280        m_sTauxDessin = rstProjSoum.Fields("TauxDessin")
1285      Else
1290        m_sTauxDessin = "0"
1295      End If

1300      If Not IsNull(rstProjSoum.Fields("TauxCoupe")) Then
1305        m_sTauxCoupe = rstProjSoum.Fields("TauxCoupe")
1310      Else
1315        m_sTauxCoupe = "0"
1320      End If

1325      If Not IsNull(rstProjSoum.Fields("TauxMachinage")) Then
1330        m_sTauxMachinage = rstProjSoum.Fields("TauxMachinage")
1335      Else
1340        m_sTauxMachinage = "0"
1345      End If

1350      If Not IsNull(rstProjSoum.Fields("TauxSoudure")) Then
1355        m_sTauxSoudure = rstProjSoum.Fields("TauxSoudure")
1360      Else
1365        m_sTauxSoudure = "0"
1370      End If

1375      If Not IsNull(rstProjSoum.Fields("TauxAssemblage")) Then
1380        m_sTauxAssemblage = rstProjSoum.Fields("TauxAssemblage")
1385      Else
1390        m_sTauxAssemblage = "0"
1395      End If

1400      If Not IsNull(rstProjSoum.Fields("TauxPeinture")) Then
1405        m_sTauxPeinture = rstProjSoum.Fields("TauxPeinture")
1410      Else
1415        m_sTauxPeinture = "0"
1420      End If

1425      If Not IsNull(rstProjSoum.Fields("TauxTest")) Then
1430        m_sTauxTest = rstProjSoum.Fields("TauxTest")
1435      Else
1440        m_sTauxTest = "0"
1445      End If

1450      If Not IsNull(rstProjSoum.Fields("TauxInstallation")) Then
1455        m_sTauxInstallation = rstProjSoum.Fields("TauxInstallation")
1460      Else
1465        m_sTauxInstallation = "0"
1470      End If

1475      If Not IsNull(rstProjSoum.Fields("TauxFormation")) Then
1480        m_sTauxFormation = rstProjSoum.Fields("TauxFormation")
1485      Else
1490        m_sTauxFormation = "0"
1495      End If

1500      If Not IsNull(rstProjSoum.Fields("TauxGestion")) Then
1505        m_sTauxGestion = rstProjSoum.Fields("TauxGestion")
1510      Else
1515        m_sTauxGestion = "0"
1520      End If

1525      If Not IsNull(rstProjSoum.Fields("TauxShipping")) Then
1530        m_sTauxShipping = rstProjSoum.Fields("TauxShipping")
1535      Else
1540        m_sTauxShipping = "0"
1545      End If

1550      If m_eType = TYPE_PROJET Then
1555        m_sTauxHebergement1 = "0"
1560        m_sTauxHebergement2 = "0"
1565        m_sTauxRepas = "0"
1570        m_sTauxTransport = "0"
1575        m_sTauxUniteMobile = "0"
1580      Else
1585        If Not IsNull(rstProjSoum.Fields("TauxHebergement1")) Then
1590          m_sTauxHebergement1 = rstProjSoum.Fields("TauxHebergement1")
1595        Else
1600          m_sTauxHebergement1 = "0"
1605        End If

1610        If Not IsNull(rstProjSoum.Fields("TauxHebergement2")) Then
1615          m_sTauxHebergement2 = rstProjSoum.Fields("TauxHebergement2")
1620        Else
1625          m_sTauxHebergement2 = "0"
1630        End If

1635        If Not IsNull(rstProjSoum.Fields("TauxRepas")) Then
1640          m_sTauxRepas = rstProjSoum.Fields("TauxRepas")
1645        Else
1650          m_sTauxRepas = "0"
1655        End If

1660        If Not IsNull(rstProjSoum.Fields("TauxTransport")) Then
1665          m_sTauxTransport = rstProjSoum.Fields("TauxTransport")
1670        Else
1675          m_sTauxTransport = "0"
1680        End If

1685        If Not IsNull(rstProjSoum.Fields("TauxUniteMobile")) Then
1690          m_sTauxUniteMobile = rstProjSoum.Fields("TauxUniteMobile")
1695        Else
1700          m_sTauxUniteMobile = "0"
1705        End If
1710      End If
       
1715      Call rstProjSoum.Close
1720      Set rstProjSoum = Nothing
1725    End If

1730    Exit Sub

AfficherErreur:

1735    woups "frmProjSoumMec", "InitialiserTempsTaux", Err, Erl
End Sub

Private Sub InitialiserTempsReel()

5       On Error GoTo AfficherErreur

10      Dim rstPunch        As ADODB.Recordset
15      Dim sDateDebut      As String
20      Dim sDateFin        As String
25      Dim sTotal          As String
30      Dim sFilterNoProjet As String

35      If Right$(txtNoProjSoum.Text, 2) = "99" Then
40        sFilterNoProjet = "LEFT(NoProjet, 6) = '" & Left$(txtNoProjSoum.Text, 6) & "'"
45      Else
50        sFilterNoProjet = "NoProjet = '" & txtNoProjSoum.Text & "'"
55      End If

60      sDateDebut = "TIMESERIAL(Left(GRB_Punch.HeureDébut,2),RIGHT(GRB_Punch.HeureDébut,2),0)"

65      sDateFin = "TIMESERIAL(Left(GRB_Punch.HeureFin,2),RIGHT(GRB_Punch.HeureFin,2),0)"

70      sTotal = "(SUM(" & sDateFin & " - " & sDateDebut & ")* 24) As Total"

75      Set rstPunch = New ADODB.Recordset

80      Call rstPunch.Open("SELECT Type, " & sTotal & " FROM GRB_Punch WHERE " & sFilterNoProjet & " And HeureFin is not null AND HeureDébut is not null GROUP BY Type", g_connData, adOpenDynamic, adLockOptimistic)

85      m_sTempsDessin = "0"
90      m_sTempsCoupe = "0"
95      m_sTempsMachinage = "0"
100     m_sTempsSoudure = "0"
105     m_sTempsAssemblage = "0"
110     m_sTempsPeinture = "0"
115     m_sTempsTest = "0"
120     m_sTempsInstallation = "0"
125     m_sTempsFormation = "0"
130     m_sTempsGestion = "0"
135     m_sTempsShipping = "0"

140     Do While Not rstPunch.EOF
145       If Not IsNull(rstPunch.Fields("Total")) Then
150         Select Case rstPunch.Fields("Type")
              Case "Dessin":       m_sTempsDessin = Round(rstPunch.Fields("Total"), 2)
155           Case "Coupe":        m_sTempsCoupe = Round(rstPunch.Fields("Total"), 2)
160           Case "Machinage":    m_sTempsMachinage = Round(rstPunch.Fields("Total"), 2)
165           Case "Soudure":      m_sTempsSoudure = Round(rstPunch.Fields("Total"), 2)
170           Case "Assemblage":   m_sTempsAssemblage = Round(rstPunch.Fields("Total"), 2)
175           Case "Peinture":     m_sTempsPeinture = Round(rstPunch.Fields("Total"), 2)
180           Case "Test":         m_sTempsTest = Round(rstPunch.Fields("Total"), 2)
185           Case "Installation": m_sTempsInstallation = Round(rstPunch.Fields("Total"), 2)
190           Case "Formation":    m_sTempsFormation = Round(rstPunch.Fields("Total"), 2)
195           Case "Gestion":      m_sTempsGestion = Round(rstPunch.Fields("Total"), 2)
200           Case "Shipping":     m_sTempsShipping = Round(rstPunch.Fields("Total"), 2)
205         End Select
210       End If

215       Call rstPunch.MoveNext
220     Loop

225     Call rstPunch.Close
230     Set rstPunch = Nothing

235     Exit Sub

AfficherErreur:

240     woups "frmChoixDateImpressionFacturation", "RemplirTempsReelsElec", Err, Erl
End Sub

Private Sub cmbProjSoum_KeyUp(KeyCode As Integer, Shift As Integer)
  
5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer
  
15      For iCompteur = 0 To cmbProjSoum.ListCount - 1
20        If UCase(cmbProjSoum.LIST(iCompteur)) = UCase(cmbProjSoum.Text) Then
25          cmbProjSoum.ListIndex = iCompteur
      
30          Exit For
35        End If
40      Next

45      Exit Sub

AfficherErreur:

50      woups "frmProjSoumMec", "cmbProjSoum_KeyUp", Err, Erl
End Sub

Private Sub cmdAjouterSection_Click()

5       On Error GoTo AfficherErreur

        'Affiche le form frmSoumissionSection
10      Call OuvrirForm(frmSoumissionSectionMec, True)
  
        'Après que l'utilisateur a refermé le form, on rafraichi le
        'contenu du combo
15      Call RemplirComboSections
  
        'On change l'ordre si elle a été changée
20      Call UpdateOrdre

25      Exit Sub

AfficherErreur:

30      woups "frmProjSoumMec", "cmdAjouterSection_Click", Err, Erl
End Sub

Private Sub cmdAnglaisFrancais_Click()

5       On Error GoTo AfficherErreur

10      If cmdAnglaisFrancais.Caption = "Anglais" Then
15        m_eLangage = ANGLAIS
    
20        cmdAnglaisFrancais.Caption = "Français"
25      Else
30        m_eLangage = FRANCAIS
    
35        cmdAnglaisFrancais.Caption = "Anglais"
40      End If

45      Call UpdateDescription
  
50      Call RemplirComboSections
    
55      Call UpdateOrdre
  
60     Exit Sub

AfficherErreur:

65     woups "frmProjSoumMec", "cmdAnglaisFrancais_Click", Err, Erl
End Sub

Private Sub UpdateDescription()

5       On Error GoTo AfficherErreur

10      Dim rstProjSoum As ADODB.Recordset
15      Dim rstPieceMec As ADODB.Recordset

20      Set rstProjSoum = New ADODB.Recordset
25      Set rstPieceMec = New ADODB.Recordset

30      If m_eType = TYPE_PROJET Then
35        Call rstProjSoum.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
40      Else
45        Call rstProjSoum.Open("SELECT * FROM GRB_Soumission_Pieces WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
50      End If

55      Do While Not rstProjSoum.EOF
60        Call rstPieceMec.Open("SELECT * FROM GRB_CatalogueMec WHERE PIECE = '" & rstProjSoum.Fields("NumItem") & "'", g_connData, adOpenDynamic, adLockOptimistic)

65        rstProjSoum.Fields("Desc_FR") = rstPieceMec.Fields("DESC_FR")
70        rstProjSoum.Fields("Desc_EN") = rstPieceMec.Fields("DESC_EN")

75        Call rstProjSoum.Update

80        Call rstPieceMec.Close

85        Call rstProjSoum.MoveNext
90      Loop

95      Set rstPieceMec = Nothing

100     Call rstProjSoum.Close
105     Set rstProjSoum = Nothing

110     Call RemplirListViewProjSoum(txtNoProjSoum.Text)

115     Exit Sub

AfficherErreur:

120     woups "frmProjSoumMec", "UpdateDescription", Err, Erl
End Sub

Private Sub cmdAnnuler_Click()

5       On Error GoTo AfficherErreur

10      frafournisseur.Visible = False
15      fraPieceTrouve.Visible = False
20      fraCommentaire.Visible = False
25      fraDateRequise.Visible = False
  
30      Screen.MousePointer = vbHourglass
 
35      Call OuvrirProjSoum(False)
  
        'Remet en mode inactif
40      Call AfficherControles(MODE_INACTIF)
  
45      m_bEnregistrement = True
  
        'Affiche l'enregistrement qui était actif avant
50      Call AfficherProjSoum(m_sAncienProjSoum)

55      m_bEnregistrement = False
    
60      m_bModeAjout = False
    
65      Screen.MousePointer = vbDefault

70      Exit Sub

AfficherErreur:

75      woups "frmProjSoumMec", "cmdAnnuler_Click", Err, Erl
End Sub

Private Sub cmdAnnulerCommentaire_Click()

5       On Error GoTo AfficherErreur

10      fraCommentaire.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMec", "cmdAnnulerCommentaire_Click", Err, Erl
End Sub

Private Sub cmdAnnulerDateRequise_Click()

5       On Error GoTo AfficherErreur

10      fraDateRequise.Visible = False

15      m_bMonthViewHasFocus = False

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumMec", "cmdAnnulerDateRequise_Click", Err, Erl
End Sub

Private Sub cmdAnnulerDateRequise_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
5       On Error GoTo AfficherErreur

10      If m_bMonthViewHasFocus = True Then
15        Call cmdAnnulerDateRequise_Click
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmProjSoumMec", "cmdAnnulerDateRequise_MouseUp", Err, Erl
End Sub

Private Sub cmdExport_Click()

5       On Error GoTo AfficherErreur

10      Dim rstProjSoum As ADODB.Recordset

15      Set rstProjSoum = New ADODB.Recordset
 
20      Call g_connData.Execute("DELETE * FROM GRB_impression_listepiece")

25      If m_eType = TYPE_PROJET Then
30        Call rstProjSoum.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
35      Else
40        Call rstProjSoum.Open("SELECT * FROM GRB_SoumissionMec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
45      End If

50      Call ExporterListePieces(rstProjSoum)

55      Call rstProjSoum.Close
60      Set rstProjSoum = Nothing

65      Exit Sub

AfficherErreur:

70      woups "frmProjSoumMec", "cmdExport_Click", Err, Erl
End Sub

Private Sub ExporterListePieces(ByVal rstProjSoum As ADODB.Recordset)

5       On Error GoTo AfficherErreur

        'Impression de la feuille de soumission
10      Dim rstPiece            As ADODB.Recordset
15      Dim rstTemp             As ADODB.Recordset
20      Dim rstImpListePiece    As ADODB.Recordset
25      Dim iCompteurPiece      As Integer
30      Dim sSousSection        As String
35      Dim sSection            As String
40      Dim sNoProjet           As String
45      Dim sNoSoumission       As String
50      Dim bAjouterSection     As Boolean
55      Dim bAjouterSousSection As Boolean
60      Dim bAjouterPiece       As Boolean
65      Dim xlsApp              As Excel.Application
70      Dim xlsWorkBook         As Excel.Workbook
75      Dim iCompteur           As Integer
80      Dim sSaveAsFileName     As String
      
85      Set rstPiece = New ADODB.Recordset
90      Set rstTemp = New ADODB.Recordset
95      Set rstImpListePiece = New ADODB.Recordset
      
100     iCompteurPiece = 1

105     Screen.MousePointer = vbHourglass
            
110     If m_eType = TYPE_PROJET Then
115       sNoProjet = rstProjSoum.Fields("IDProjet")
120       sNoSoumission = rstProjSoum.Fields("IDSoumission")

125       Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoProjet & "' And Type = 'M'", g_connData, adOpenDynamic, adLockOptimistic)
130     Else
135       sNoProjet = vbNullString
140       sNoSoumission = rstProjSoum.Fields("IDSoumission")

145       Call rstPiece.Open("SELECT * FROM GRB_Soumission_Pieces WHERE IDSoumission = '" & sNoSoumission & "' And Type = 'M'", g_connData, adOpenDynamic, adLockOptimistic)
150     End If
    
155     Do While Not rstPiece.EOF
160       If rstPiece.Fields("Visible") = True Then
165         bAjouterSection = True
170         bAjouterSousSection = True
175         bAjouterPiece = True

180         rstImpListePiece.CursorLocation = adUseClient

185         rstImpListePiece.Filter = ""

190         Call rstImpListePiece.Open("SELECT * FROM GRB_Impression_ListePiece WHERE IDSection = '" & rstPiece.Fields("IDSection") & "'", g_connData, adOpenDynamic, adLockOptimistic)

195         If Not rstImpListePiece.EOF Then
200           bAjouterSection = False

205           Do While Not rstImpListePiece.EOF
210             If rstImpListePiece.Fields("NomSousSection") = rstPiece.Fields("SousSection") Then
215               bAjouterSousSection = False

220               If rstPiece.Fields("NumItem") <> "Texte" And rstPiece.Fields("NumItem") <> "Text" Then
225                 If rstImpListePiece.Fields("NumItem") = rstPiece.Fields("NumItem") Then
230                   bAjouterPiece = False

235                   rstImpListePiece.Fields("Qté") = CDbl(Replace(rstImpListePiece.Fields("Qté"), ".", ",")) + CDbl(rstPiece.Fields("Qté"))

240                   Call rstImpListePiece.Update

245                   If rstImpListePiece.Fields("Qté") = 0 Then
250                     Call rstImpListePiece.Delete

255                     rstImpListePiece.Filter = "NomSousSection = '" & Replace(rstPiece.Fields("SousSection"), "'", "''") & "'"

260                     If rstImpListePiece.RecordCount = 1 Then
265                       Call rstImpListePiece.Delete

270                       rstImpListePiece.Filter = "IDSection = '" & rstPiece.Fields("IDSection") & "'"

275                       If rstImpListePiece.RecordCount = 1 Then
280                         Call rstImpListePiece.Delete
285                       End If
290                     End If
295                   End If
300                 End If
305               Else
310                 Exit Do
315               End If
320             End If

325             Call rstImpListePiece.MoveNext
330           Loop
335         End If

340         If bAjouterSection = True Then
345           If m_eLangage = ANGLAIS Then
350             sSection = "NomSectionEN"
355           Else
360             sSection = "NomSectionFR"
365           End If

370           Call rstTemp.Open("SELECT " & sSection & " FROM GRB_SoumProjSectionMec WHERE IDSection = " & rstPiece.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
          
375           Call rstImpListePiece.AddNew
          
380           rstImpListePiece("NoLigne") = iCompteurPiece
385           rstImpListePiece("IDSoumission") = sNoSoumission
          
390           If Not IsNull(rstTemp.Fields(sSection)) Then
395             rstImpListePiece.Fields("Section") = rstTemp.Fields(sSection)
400           Else
405             rstImpListePiece.Fields("Section") = " "
410           End If

415           rstImpListePiece.Fields("IDSection") = rstPiece.Fields("IDSection")
          
420           Call rstImpListePiece.Update
                   
425           iCompteurPiece = iCompteurPiece + 1
          
430           Call rstTemp.Close
435         End If
          
440         If bAjouterSousSection = True Then
445           sSousSection = rstPiece.Fields("SousSection")
          
450           If sSousSection = S_PAS_SOUS_SECTION Then
455             sSousSection = " "
460           End If
          
465           Call rstImpListePiece.AddNew
          
470           rstImpListePiece.Fields("NoLigne") = iCompteurPiece
475           rstImpListePiece.Fields("IDSoumission") = sNoSoumission
480           rstImpListePiece.Fields("SousSection") = sSousSection
485           rstImpListePiece.Fields("NomSousSection") = rstPiece.Fields("SousSection")
490           rstImpListePiece.Fields("IDSection") = rstPiece.Fields("IDSection")
          
495           Call rstImpListePiece.Update
          
500           iCompteurPiece = iCompteurPiece + 1
505         End If
              
510         If bAjouterPiece = True Then
              'ajoute une piece dans la liste de pièce
515           Call rstImpListePiece.AddNew
      
520           rstImpListePiece.Fields("NoLigne") = iCompteurPiece
525           rstImpListePiece.Fields("IDSoumission") = sNoSoumission
530           rstImpListePiece.Fields("numitem") = rstPiece.Fields("numitem")
535           rstImpListePiece.Fields("qté") = rstPiece.Fields("qté")
       
540           If m_eLangage = ANGLAIS Then
545             rstImpListePiece.Fields("description") = rstPiece.Fields("DESC_EN")
550           Else
555             rstImpListePiece.Fields("description") = rstPiece.Fields("DESC_FR")
560           End If
        
565           rstImpListePiece.Fields("manufact") = rstPiece.Fields("manufact")

570           rstImpListePiece.Fields("NomSousSection") = rstPiece.Fields("SousSection")
575           rstImpListePiece.Fields("IDSection") = rstPiece.Fields("IDSection")
          
580           Call rstImpListePiece.Update
        
585           iCompteurPiece = iCompteurPiece + 1
590         End If

595         Call rstImpListePiece.Close
600       End If
 
605       Call rstPiece.MoveNext
610     Loop
   
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        ' rapport liste piece, met dans l'ordre de ligne '
        ''''''''''''''''''''''''''''''''''''''''''''''''''
615     rstImpListePiece.CursorLocation = adUseClient

620     Call rstImpListePiece.Open("SELECT * FROM GRB_impression_Listepiece WHERE IDSoumission = '" & sNoSoumission & "' ORDER BY NoLigne", g_connData, adOpenDynamic, adLockOptimistic)
        
625     Set xlsApp = New Excel.Application

630     Set xlsWorkBook = xlsApp.Workbooks.Add

635     xlsApp.Range("A1") = "Liste de matériel ( " & txtNoProjSoum.Text & " ) "
640     xlsApp.Range("A1").Font.Bold = True
645     xlsApp.Range("A1").Font.Underline = xlUnderlineStyleSingle
650     xlsApp.Range("A1").HorizontalAlignment = xlCenter
655     xlsApp.Range("A1").Font.SIZE = 14

660     Call xlsApp.Range("A1:D1").Merge
        
665     xlsApp.Range("A4") = "Qté"
670     xlsApp.Range("A4").Font.Bold = True
675     xlsApp.Range("A4").HorizontalAlignment = xlCenter

680     xlsApp.Range("B4") = "No. Item"
685     xlsApp.Range("B4").Font.Bold = True
690     xlsApp.Range("B4").HorizontalAlignment = xlCenter

695     xlsApp.Range("C4") = "Description"
700     xlsApp.Range("C4").Font.Bold = True
705     xlsApp.Range("C4").HorizontalAlignment = xlCenter

710     xlsApp.Range("D4") = "Manufacturier"
715     xlsApp.Range("D4").Font.Bold = True
720     xlsApp.Range("D4").HorizontalAlignment = xlCenter

725     xlsApp.Range("A4:D4").Borders(xlEdgeBottom).LineStyle = xlContinuous
730     xlsApp.Range("A4:D4").Borders(xlEdgeBottom).Weight = xlMedium
735     xlsApp.Range("A4:D4").Borders(xlEdgeBottom).ColorIndex = xlAutomatic

740     xlsApp.Range("A4:D4").Borders(xlInsideVertical).LineStyle = xlContinuous
745     xlsApp.Range("A4:D4").Borders(xlInsideVertical).Weight = xlMedium
750     xlsApp.Range("A4:D4").Borders(xlInsideVertical).ColorIndex = xlAutomatic

755     iCompteur = 5

760     Do While Not rstImpListePiece.EOF
765       xlsApp.Range("A" & iCompteur) = rstImpListePiece.Fields("Qté")

770       If IsNull(rstImpListePiece.Fields("Section")) Then
775         xlsApp.Range("B" & iCompteur) = rstImpListePiece.Fields("NumItem")
780       Else
785         xlsApp.Range("B" & iCompteur) = rstImpListePiece.Fields("Section")
790         xlsApp.Range("B" & iCompteur).Font.Bold = True
795       End If

800       If IsNull(rstImpListePiece.Fields("SousSection")) Then
805         xlsApp.Range("C" & iCompteur) = rstImpListePiece.Fields("Description")
810       Else
815         xlsApp.Range("C" & iCompteur) = rstImpListePiece.Fields("SousSection")
820         xlsApp.Range("C" & iCompteur).Font.Bold = True
825       End If

830       xlsApp.Range("D" & iCompteur) = rstImpListePiece.Fields("Manufact")

835       xlsApp.Range("A" & iCompteur & ":D" & iCompteur).Borders(xlEdgeBottom).LineStyle = xlContinuous
840       xlsApp.Range("A" & iCompteur & ":D" & iCompteur).Borders(xlEdgeBottom).Weight = xlThin
845       xlsApp.Range("A" & iCompteur & ":D" & iCompteur).Borders(xlEdgeBottom).ColorIndex = xlAutomatic

850       xlsApp.Range("A" & iCompteur & ":D" & iCompteur).Borders(xlInsideVertical).LineStyle = xlContinuous
855       xlsApp.Range("A" & iCompteur & ":D" & iCompteur).Borders(xlInsideVertical).Weight = xlThin
860       xlsApp.Range("A" & iCompteur & ":D" & iCompteur).Borders(xlInsideVertical).ColorIndex = xlAutomatic

865       Call rstImpListePiece.MoveNext

870       iCompteur = iCompteur + 1
875     Loop

880     iCompteur = iCompteur - 1

885     xlsApp.Range("A4:D" & iCompteur).Borders(xlEdgeBottom).LineStyle = xlContinuous
890     xlsApp.Range("A4:D" & iCompteur).Borders(xlEdgeBottom).Weight = xlMedium
895     xlsApp.Range("A4:D" & iCompteur).Borders(xlEdgeBottom).ColorIndex = xlAutomatic

900     xlsApp.Range("A4:D" & iCompteur).Borders(xlEdgeTop).LineStyle = xlContinuous
905     xlsApp.Range("A4:D" & iCompteur).Borders(xlEdgeTop).Weight = xlMedium
910     xlsApp.Range("A4:D" & iCompteur).Borders(xlEdgeTop).ColorIndex = xlAutomatic

915     xlsApp.Range("A4:D" & iCompteur).Borders(xlEdgeLeft).LineStyle = xlContinuous
920     xlsApp.Range("A4:D" & iCompteur).Borders(xlEdgeLeft).Weight = xlMedium
925     xlsApp.Range("A4:D" & iCompteur).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

930     xlsApp.Range("A4:D" & iCompteur).Borders(xlEdgeRight).LineStyle = xlContinuous
935     xlsApp.Range("A4:D" & iCompteur).Borders(xlEdgeRight).Weight = xlMedium
940     xlsApp.Range("A4:D" & iCompteur).Borders(xlEdgeRight).ColorIndex = xlAutomatic

945     Call xlsApp.Columns("A:A").EntireColumn.AutoFit
950     Call xlsApp.Columns("B:B").EntireColumn.AutoFit
955     Call xlsApp.Columns("C:C").EntireColumn.AutoFit
960     Call xlsApp.Columns("D:D").EntireColumn.AutoFit

965     Call rstImpListePiece.Close
970     Set rstImpListePiece = Nothing

975     Screen.MousePointer = vbDefault

980     sSaveAsFileName = xlsApp.GetSaveAsFilename(txtNoProjSoum.Text & ".xls", "Fichiers Excel (*.xlx), *.xls")

985     If sSaveAsFileName <> "Faux" Then
990       Call xlsWorkBook.SaveAs(sSaveAsFileName)
995     End If

1000    xlsWorkBook.Saved = True
        
1005    Call xlsWorkBook.Close

1010    Set xlsWorkBook = Nothing

1015    Call xlsApp.Quit

1020    Set xlsApp = Nothing

1025    Set rstTemp = Nothing

1030    Exit Sub

AfficherErreur:

1035    woups "frmProjSoumMec", "ExporterListePieces", Err, Erl
End Sub

Private Sub cmdOKCommentaire_Click()

5       On Error GoTo AfficherErreur

10      lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_COMMENTAIRE) = txtcommentaire.Text

15      fraCommentaire.Visible = False

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumMec", "cmdOKCommentaire_Click", Err, Erl
End Sub

Private Sub cmdOKDateRequise_Click()

5       On Error GoTo AfficherErreur

10      Dim datDate As Date

15      datDate = DateSerial(mvwDateRequise.Year, mvwDateRequise.Month, mvwDateRequise.Day)

20      lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_DATE_REQUISE) = ConvertDate(datDate)

25      lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_DATE_REQUISE).ForeColor = COLOR_ORANGE

30      fraDateRequise.Visible = False

35      m_bMonthViewHasFocus = False

40      Exit Sub

AfficherErreur:

45      woups "frmProjSoumMec", "cmdOKDateRequise_Click", Err, Erl
End Sub

Private Sub cmdOKDateRequise_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
5       On Error GoTo AfficherErreur

10      If m_bMonthViewHasFocus = True Then
15        Call cmdOKDateRequise_Click
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmProjSoumMec", "cmdOKDateRequise_MouseUp", Err, Erl
End Sub

Private Sub cmdAnnulerFRS_Click()

5       On Error GoTo AfficherErreur

10      frafournisseur.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMec", "cmdAnnulerFRS_Click", Err, Erl
End Sub

Private Sub cmdAnnulerPrix_Click()

5       On Error GoTo AfficherErreur

10      fraPrix.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMec", "cmdAnnulerPrix_Click", Err, Erl
End Sub

Private Sub cmdBavards_Click()

5       On Error GoTo AfficherErreur

10      Call RemplirListViewSuppression

15      lvwBavard.Visible = True

20      Call lvwBavard.SetFocus

25      Exit Sub

AfficherErreur:

30      woups "frmProjSoumMec", "cmdBavards_Click", Err, Erl
End Sub

Private Sub cmdBonCommande_Click()

5       On Error GoTo AfficherErreur

10      Dim rstProjSoum As ADODB.Recordset
15      Dim sUser       As String

20      If Right$(txtNoProjSoum.Text, 2) = "99" Then
25        Call MsgBox("Vous ne pouvez pas commander de pièce à partir de ce projet!", vbOKOnly, "Erreur")

30        Exit Sub
35      End If

40      Set rstProjSoum = New ADODB.Recordset

45      Call rstProjSoum.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

50      If rstProjSoum.Fields("Ouvert") = True And rstProjSoum.Fields("Verrouillé") = False Then
55        If VerifierSiOuvert(sUser) = False Then
60          If lvwSoumission.ListItems.count > 0 Then
65            Call frmChoixBonCommande.Afficher(txtNoProjSoum.Text, Me, m_eLangage)
70          Else
75            Call MsgBox("Il n'y a pas de pièces à commander pour ce projet!", vbOKOnly, "Erreur")
80          End If
85        Else
90          Call MsgBox("Ce projet est en modification par " & sUser & "!", vbOKOnly, "Erreur")
95        End If
100     Else
105       If rstProjSoum.Fields("Ouvert") = False Then
110         Call MsgBox("Ce projet est fermé!", vbOKOnly, "Erreur")
115       Else
120         Call MsgBox("Ce projet est verrouillé!", vbOKOnly, "Erreur")
125       End If
130     End If
        
135     Call rstProjSoum.Close
140     Set rstProjSoum = Nothing

145     If m_bModifFournisseurBC = True Then
150       Call cmbProjSoum_Click
155     End If

160     Exit Sub

AfficherErreur:

165     woups "frmProjSoumMec", "cmdBonCommande_Click", Err, Erl
End Sub

Public Sub Commande()

5       On Error GoTo AfficherErreur

10      Dim rstProjet    As ADODB.Recordset
15      Dim rstPiece     As ADODB.Recordset
20      Dim rstBCPiece   As ADODB.Recordset
25      Dim rstBC        As ADODB.Recordset
30      Dim rstFRS       As ADODB.Recordset
35      Dim iIDFRS       As Integer
40      Dim sFRS         As String
45      Dim sNoBC        As String
50      Dim sWhere       As String
55      Dim sDateRequise As String
60      Dim sNoLigne     As String
65      Dim bPremier     As Boolean

70      Set rstProjet = New ADODB.Recordset

75      Call rstProjet.Open("SELECT ProchaineCommande FROM GRB_ProjetMec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

80      If Not IsNull(rstProjet.Fields("ProchaineCommande")) Then
85        rstProjet.Fields("ProchaineCommande") = rstProjet.Fields("ProchaineCommande") + 1

90        Call rstProjet.Update
95      End If

100     Call rstProjet.Close
105     Set rstProjet = Nothing

110     sFRS = DR_Commande.Sections("Section2").Controls("lblFournisseur").Caption
115     sNoBC = DR_Commande.Sections("Section2").Controls("lblNoBC").Caption

120     Set rstBC = New ADODB.Recordset
125     Set rstFRS = New ADODB.Recordset

130     Call rstBC.Open("SELECT * FROM GRB_BonsCommandes WHERE NoBonCommande = '" & sNoBC & "'", g_connData, adOpenDynamic, adLockOptimistic)
        
135     Do While Not rstBC.EOF
140       Call rstFRS.Open("SELECT IDFRS, NomFournisseur FROM GRB_Fournisseur WHERE IDFRS = " & rstBC.Fields("NoFournisseur"), g_connData, adOpenDynamic, adLockOptimistic)

145       If rstFRS.Fields("NomFournisseur") = sFRS Then
150         iIDFRS = rstFRS.Fields("IDFRS")

155         sDateRequise = rstBC.Fields("DateRequise")

160         Call rstFRS.Close

165         Exit Do
170       End If

175       Call rstFRS.Close

180       Call rstBC.MoveNext
185     Loop

190     Set rstFRS = Nothing

195     Call rstBC.Close
200     Set rstBC = Nothing
        
        'Ouverture du recordset du Bon de commande pour savoir quelles pièces
        'ont été commandées
205     Set rstBCPiece = New ADODB.Recordset
        
210     Call rstBCPiece.Open("SELECT NoItem, NuméroLigne FROM GRB_BonsCommandes_Pieces WHERE NoFournisseur = " & iIDFRS & " AND NoBonCommande = '" & sNoBC & "'", g_connData, adOpenDynamic, adLockOptimistic)
                
        'Tant que ce n'est pas la fin des enregistrements
215     Do While Not rstBCPiece.EOF
220       bPremier = True

225       If Not IsNull(rstBCPiece.Fields("NoItem")) Then
230         sNoLigne = rstBCPiece.Fields("NuméroLigne")

235         If sWhere = vbNullString Then
240           sWhere = "(IDProjet = '" & txtNoProjSoum.Text & "')"

245           If InStr(1, sNoLigne, ",") = 0 Then
250             sWhere = sWhere & " AND ((NumItem = '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "' AND NuméroLigne = " & rstBCPiece.Fields("NuméroLigne") & ")"
255           Else
260             Do While InStr(1, sNoLigne, ",") > 0
265               If bPremier = True Then
270                 sWhere = sWhere & " AND ((NumItem = '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "' AND NuméroLigne = " & Left$(sNoLigne, InStr(1, sNoLigne, ",") - 1) & ")"

275                 bPremier = False
280               Else
285                 sWhere = sWhere & " OR (NumItem = '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "' AND NuméroLigne = " & Left$(sNoLigne, InStr(1, sNoLigne, ",") - 1) & ")"
290               End If

295               sNoLigne = Right$(sNoLigne, Len(sNoLigne) - (InStr(1, sNoLigne, ",") + 1))
300             Loop

305             If Trim$(sNoLigne) <> "" Then
310               sWhere = sWhere & " OR (NumItem = '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "' AND NuméroLigne = " & sNoLigne & ")"
315             End If
320           End If
325         Else
330           If InStr(1, sNoLigne, ",") = 0 Then
335             sWhere = sWhere & " OR (NumItem = '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "' AND NuméroLigne = " & rstBCPiece.Fields("NuméroLigne") & ")"
340           Else
345             Do While InStr(1, sNoLigne, ",") > 0
350               sWhere = sWhere & " OR (NumItem = '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "' AND NuméroLigne = " & Left$(sNoLigne, InStr(1, sNoLigne, ",") - 1) & ")"

355               sNoLigne = Right$(sNoLigne, Len(sNoLigne) - (InStr(1, sNoLigne, ",") + 1))
360             Loop

365             If Trim$(sNoLigne) <> "" Then
370               sWhere = sWhere & " OR (NumItem = '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "' AND NuméroLigne = " & sNoLigne & ")"
375             End If
380           End If
385         End If
390       End If
    
395       Call rstBCPiece.MoveNext
400     Loop

405     sWhere = sWhere & ")"
  
410     Call rstBCPiece.Close
415     Set rstBCPiece = Nothing
  
420     Set rstPiece = New ADODB.Recordset
  
425     Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
  
430     Do While Not rstPiece.EOF
435       rstPiece.Fields("Commandé") = True

440       rstPiece.Fields("DateCommande") = ConvertDate(Date)

445       rstPiece.Fields("DateRequise") = sDateRequise

450       rstPiece.Fields("NomCommande") = g_sEmploye

455       rstPiece.Fields("NoSéquentiel") = Right$(sNoBC, 3)
    
460       Call rstPiece.Update
    
465       Call rstPiece.MoveNext
470     Loop
  
475     Call rstPiece.Close
480     Set rstPiece = Nothing
          
485     Call RemplirListViewProjSoum(txtNoProjSoum.Text)

490     Exit Sub

AfficherErreur:

495     woups "frmProjSoumMec", "Commande", Err, Erl
End Sub

Private Sub cmdCatalogue_Click()

5       On Error GoTo AfficherErreur

        'Pour ouvrir le catalogue mécanique
10      Screen.MousePointer = vbHourglass
 
15      Call FrmCatalogueMec.AfficherForm(cmbPieces.Text, "", "")

20      Screen.MousePointer = vbDefault
  
25      Exit Sub

AfficherErreur:

30      woups "frmProjSoumMec", "cmdCatalogue_Click", Err, Erl
End Sub

Private Sub cmdMauvaisPrix_Click()

5       On Error GoTo AfficherErreur

10      Call MauvaisPrix

15      Exit Sub

AfficherErreur:

20       woups "frmProjSoumMec", "cmdMauvaisPrix_Click", Err, Erl
End Sub

Private Sub MauvaisPrix()

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer

15      If lvwSoumission.ListItems.count > 0 Then
20        If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PROFIT).Tag <> "EXTRA" Then
            'Si ce n'est pas une section
25          If lvwSoumission.SelectedItem.Tag <> vbNullString Then
              'Si ce n'est pas une sous-section
30            If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) <> vbNullString Then
                'Si ce n'est pas du texte
35              If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) <> "Texte" And lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) <> "Text" Then
                  'Si la quantité est plus grande que 0
40                If CDbl(Replace(lvwSoumission.SelectedItem.Text, "*", vbNullString)) > 0 Then
45                  Call ViderChamps_frs

50                  Call RemplirComboFournisseur

55                  For iCompteur = 0 To cmbfrs.ListCount - 1
60                    If cmbfrs.ItemData(iCompteur) = lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_DISTRIB).Tag Then
65                      cmbfrs.ListIndex = iCompteur

70                      Exit For
75                    End If
80                  Next

85                  cmbfrs.Locked = True

90                  fraPrix.Tag = lvwSoumission.SelectedItem.Index

95                  m_bMauvaisPrix = True

100                 fraPrix.Visible = True

105                 Call txtPrixList.SetFocus
110               Else
115                 Call MsgBox("La quantité est déjà dans le négatif!", vbOKOnly, "Erreur")
120               End If
125             End If
130           End If
135         End If
140       Else
145         Call MsgBox("Cette commande doit être faite dans le projet " & lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PROVENANCE), vbOKOnly, "Erreur")
150       End If
155     End If

160     Exit Sub

AfficherErreur:

165     woups "frmProjSoumMec", "MauvaisPrix", Err, Erl
End Sub

Private Sub cmdMaterielInutile_Click()

5       On Error GoTo AfficherErreur

10      Call MaterielInutile

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMec", "cmdMaterielInutile_Click", Err, Erl
End Sub

Private Sub MaterielInutile()

5       On Error GoTo AfficherErreur

10      Dim itmProjet As ListItem

15      If lvwSoumission.ListItems.count > 0 Then
20        Set itmProjet = lvwSoumission.SelectedItem

25        If itmProjet.ListSubItems(I_COL_SOUM_PIECE).ForeColor <> COLOR_ROSE And itmProjet.ListSubItems(I_COL_SOUM_PIECE).ForeColor <> COLOR_BLEU Then
            'Si ce n'est pas une section
30          If itmProjet.Tag <> vbNullString Then
              'Si ce n'est pas une sous-section
35            If itmProjet.SubItems(I_COL_SOUM_PIECE) <> vbNullString Then
                'Si ce n'est pas du texte
40              If itmProjet.SubItems(I_COL_SOUM_PIECE) <> "Texte" And itmProjet.SubItems(I_COL_SOUM_PIECE) <> "Text" Then
                  'Si la quantité est plus grande que 0
45                If CDbl(Replace(itmProjet.Text, "*", vbNullString)) > 0 Then
50                  m_bPieceInutile = True
55                  m_bRecherchePiece = False
60                  m_bChangementFRS = False

65                  Call AfficherListeFournisseurs

70                  If lvwfournisseur.ListItems.count = 0 Then
75                    Call MsgBox("Il n'y a aucun fournisseur pour cette pièce!", vbOKOnly, "Erreur")
 
80                    Exit Sub
85                  Else
90                    frafournisseur.Visible = True
95                  End If
100               Else
105                 Call MsgBox("La quantité est déjà dans le négatif!", vbOKOnly, "Erreur")
110               End If
115             End If
120           End If
125         End If
130       Else
135         Call MsgBox("Cette commande doit être faite dans le projet " & itmProjet.SubItems(I_COL_SOUM_PROVENANCE), vbOKOnly, "Erreur")
140       End If
145     End If

150     Exit Sub

AfficherErreur:

155     woups "frmProjSoumMec", "MaterielInutile", Err, Erl
End Sub

Private Sub cmdCopier_Click()

5       On Error GoTo AfficherErreur

10      Dim rstProjSoum  As ADODB.Recordset
15      Dim sNoProjSoum  As String
20      Dim sUser        As String
25      Dim bExiste      As Boolean
30      Dim bVariables   As Boolean
35      Dim bTauxHoraire As Boolean
40      Dim bPrixPieces  As Boolean
45      Dim bNoValide    As Boolean
  
        'Si le combo n'est pas vide
50      If cmbProjSoum.ListCount > 0 Then
55        If VerifierSiOuvert(sUser) = False Then
            'Demande du numéro
60          sNoProjSoum = InputBox("Quel est le numéro de la soumission?")
  
65          If Trim$(sNoProjSoum) <> vbNullString Then
70            Screen.MousePointer = vbHourglass
  
75            bNoValide = True

80            If ValiderFormatNumeroProjSoum(sNoProjSoum) = False Then
85              bNoValide = False
90            End If

95            If bNoValide = True Then
100             If ValiderFormatMecanique(sNoProjSoum) = False Then
105               bNoValide = False
110             End If
115           End If

120           If bNoValide = True Then
125             If ValiderFormatSoumission(sNoProjSoum) = False Then
130               bNoValide = False
135             End If
140           End If

145           If bNoValide = False Then
150             Screen.MousePointer = vbDefault

155             Exit Sub
160           End If

165           sNoProjSoum = UCase(sNoProjSoum)
    
170           Set rstProjSoum = New ADODB.Recordset
  
175           Call rstProjSoum.Open("SELECT * FROM GRB_SoumissionElec WHERE IDSoumission = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)
        
180           If rstProjSoum.EOF Then
185             bExiste = False
190           Else
195             bExiste = True

200             Call MsgBox("Ce numéro existe dans les soumissions électriques!", vbOKOnly, "Erreur")
205           End If

210           Call rstProjSoum.Close

215           If bExiste = False Then
220             Call rstProjSoum.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)
          
225             If rstProjSoum.EOF Then
230               bExiste = False
235             Else
240               bExiste = True
  
245               Call MsgBox("Ce numéro existe dans les projets électriques!", vbOKOnly, "Erreur")
250             End If
  
255             Call rstProjSoum.Close
260           End If

265           If bExiste = False Then
270             Call rstProjSoum.Open("SELECT * FROM GRB_SoumissionMec WHERE IDSoumission = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)
  
275             If rstProjSoum.EOF Then
280               bExiste = False
285             Else
290               bExiste = True
  
295               Call MsgBox("Ce numéro existe dans les soumissions mécaniques!", vbOKOnly, "Erreur")
300             End If

305             Call rstProjSoum.Close
310           End If
          
315           If bExiste = False Then
320             Call rstProjSoum.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)
  
325             If rstProjSoum.EOF Then
330               bExiste = False
335             Else
340               bExiste = True
  
345               Call MsgBox("Ce numéro existe dans les projets mécaniques!", vbOKOnly, "Erreur")
350             End If

355             Call rstProjSoum.Close
360           End If
      
              'Si il n'existe pas, on l'ajoute
365           If bExiste = False Then
370             Call rstProjSoum.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)

375             If Not rstProjSoum.EOF Then
380               If rstProjSoum.Fields("Ouvert") = False Then
385                 Call MsgBox("Ce numéro est fermé!", vbOKOnly, "Erreur")

390                 Call rstProjSoum.Close
395                 Set rstProjSoum = Nothing

400                 Screen.MousePointer = vbDefault

405                 Exit Sub
410               End If
415             End If

420             Call rstProjSoum.Close
425             Set rstProjSoum = Nothing

430             If MsgBox("Voulez-vous mettre à jour les variables systèmes?" & vbNewLine & _
                          "-  % Profit" & vbNewLine & _
                          "-  % Commission" & vbNewLine & _
                          "-  % Imprévu", vbYesNo) = vbYes Then
435               bVariables = True
440             Else
445               bVariables = False
450             End If

455             If MsgBox("Voulez-vous mettre à jour les taux horaires?", vbYesNo) = vbYes Then
460               bTauxHoraire = True
465             Else
470               bTauxHoraire = False
475             End If

480             If MsgBox("Voulez-vous mettre à jour le prix des pièces?", vbYesNo) = vbYes Then
485               bPrixPieces = True
490             Else
495               bPrixPieces = False
500             End If

505             m_bModeAjout = True
510             m_bModeAffichage = False
        
515             m_bTempsDejaOuvert = True
             
520             If bVariables = True Then
                  'On ré-initialise les variables
525               Call InitialiserVariables(sNoProjSoum)
530             End If

535             If bTauxHoraire = True Then
540               Call InitialiserNouveauxTaux
545             End If
        
                'Rapetisse le listview de la soumission pour afficher le lvwPiece
550             lvwSoumission.Height = lvwSoumission.Height * 0.49
555             lvwSoumission.Top = lvwPieces.Top + lvwPieces.Height + (lvwSoumission.Height * 0.02)
        
                'On met en mode modif
560             Call AfficherControles(MODE_AJOUT_MODIF)
        
565             If bPrixPieces = True Then
                  'On recalcul le prix des pièces
570               Call UpdatePieces
575             End If
        
580             Call UpdateOrdre
        
585             If bVariables = True Or bTauxHoraire = True Or bPrixPieces = True Then
                  'On recalcul le prix total
590               Call CalculerPrix
595             End If
        
600             Call BarrerChamps(False)
        
605             txtNoProjSoum.Text = sNoProjSoum
610             txtNoSoumission.Text = vbNullString
615           End If
      
620           Screen.MousePointer = vbDefault
625         End If
630       Else
635         If m_eType = TYPE_PROJET Then
640           Call MsgBox("Ce projet est en modification par " & sUser & "!", vbOKOnly, "Erreur")
645         Else
650           Call MsgBox("Cette soumission est en modification par " & sUser & "!", vbOKOnly, "Erreur")
655         End If
660       End If
665     End If

670     Exit Sub

AfficherErreur:

675     woups "frmProjSoumMec", "cmdCopier_Click", Err, Erl
End Sub

Private Sub UpdateOrdre()

5       On Error GoTo AfficherErreur

        'Cette procédure sert à changer l'ordre des sections dans la soumission
10      Dim rstOrdre     As ADODB.Recordset
15      Dim rstCount     As ADODB.Recordset
20      Dim rstSection   As ADODB.Recordset
25      Dim iCompteur    As Integer
30      Dim iCompteur2   As Integer
35      Dim iIndexCopie  As Integer
40      Dim iSection     As Integer
45      Dim iIndex       As Integer
50      Dim iNbreSection As Integer
55      Dim bPremier     As Boolean
60      Dim itmProjSoum  As ListItem
65      Dim sSection     As String
  
        'Boucle pour changer la valeur de l'ordre dans le ListItem
70      Set rstOrdre = New ADODB.Recordset
        
75      For iCompteur = 1 To lvwSoumission.ListItems.count
          'Si ce n'est pas une section
80        If lvwSoumission.ListItems(iCompteur).Tag <> vbNullString Then
85          Call rstOrdre.Open("SELECT Ordre FROM GRB_SoumProjSectionMec WHERE IDSection = " & lvwSoumission.ListItems(iCompteur).Tag, g_connData, adOpenDynamic, adLockOptimistic)
        
90          lvwSoumission.ListItems(iCompteur).ListSubItems(I_COL_SOUM_MANUFACT).Tag = rstOrdre.Fields("ordre")
       
95          Call rstOrdre.Close
100       End If
105     Next
    
110     Set rstOrdre = Nothing
    
115     Set rstCount = New ADODB.Recordset
    
120     Call rstCount.Open("SELECT COUNT(IDSection) as NbreSection FROM GRB_SoumProjSectionMec", g_connData, adOpenDynamic, adLockOptimistic)
  
125     iNbreSection = rstCount.Fields("NbreSection")
    
130     Call rstCount.Close
135     Set rstCount = Nothing
     
        'Il faut enlever les sections car ils n'ont pas d'ordre et il ne font
        'que nuire
140     For iCompteur = 1 To lvwSoumission.ListItems.count
145       If lvwSoumission.ListItems(iCompteur - iSection).Tag = vbNullString Then
150         Call lvwSoumission.ListItems.Remove(iCompteur - iSection)
        
155         iSection = iSection + 1
160       End If
165     Next
    
170     iIndex = 1
        
175     Set rstSection = New ADODB.Recordset
        
        'Boucle pour replacer le ListItem à la bonne place
180     For iCompteur = 1 To iNbreSection
185       bPremier = True
      
190       iCompteur2 = iIndex
      
195       Do While iCompteur2 <= lvwSoumission.ListItems.count
            'Si le tag est la premiere ordre
200         If lvwSoumission.ListItems(iCompteur2).ListSubItems(I_COL_SOUM_MANUFACT).Tag = iCompteur Then
              'Si la première fois qu'on trouve cette ordre
205           If bPremier = True Then
                'on ajoute la section
210             Set itmProjSoum = lvwSoumission.ListItems.Add(iIndex)
            
215             Call ValeurParDefaut(itmProjSoum)
             
220             If m_eLangage = ANGLAIS Then
225               sSection = "NomSectionEN"
230             Else
235               sSection = "NomSectionFR"
240             End If
             
245             Call rstSection.Open("SELECT " & sSection & " FROM GRB_SoumProjSectionMec WHERE IDSection = " & lvwSoumission.ListItems(iCompteur2 + 1).Tag, g_connData, adOpenDynamic, adLockOptimistic)
              
250             If Not IsNull(rstSection.Fields(sSection)) Then
255               itmProjSoum.SubItems(I_COL_SOUM_PIECE) = rstSection.Fields(sSection)
260             Else
265               itmProjSoum.SubItems(I_COL_SOUM_PIECE) = vbNullString
270             End If
          
275             itmProjSoum.ListSubItems(I_COL_SOUM_PIECE).Bold = True
              
280             Call rstSection.Close
                        
285             iIndex = iIndex + 1
290             iCompteur2 = iCompteur2 + 1
          
295             bPremier = False
300           End If
            
              'On ajoute la pièce
305           Set itmProjSoum = lvwSoumission.ListItems.Add(iIndex)
           
310           iIndexCopie = iCompteur2 + 1

315           itmProjSoum.Checked = lvwSoumission.ListItems(iIndexCopie).Checked
          
320           itmProjSoum.Text = lvwSoumission.ListItems(iIndexCopie).Text

325           itmProjSoum.ForeColor = lvwSoumission.ListItems(iIndexCopie).ForeColor

330           itmProjSoum.Tag = lvwSoumission.ListItems(iIndexCopie).Tag

335           itmProjSoum.Bold = lvwSoumission.ListItems(iIndexCopie).Bold
          
340           itmProjSoum.SubItems(I_COL_SOUM_PIECE) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_PIECE)
345           itmProjSoum.ListSubItems(I_COL_SOUM_PIECE).Tag = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_PIECE).Tag

350           itmProjSoum.ListSubItems(I_COL_SOUM_PIECE).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_PIECE).ForeColor
          
355           itmProjSoum.ListSubItems(I_COL_SOUM_PIECE).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_PIECE).Bold
         
360           itmProjSoum.SubItems(I_COL_SOUM_DESCR) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_DESCR)
365           itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_DESCR).Tag

370           itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_DESCR).ForeColor
          
375           itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_DESCR).Bold = True
          
380           itmProjSoum.SubItems(I_COL_SOUM_MANUFACT) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_MANUFACT)
385           itmProjSoum.ListSubItems(I_COL_SOUM_MANUFACT).Tag = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_MANUFACT).Tag

390           itmProjSoum.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_MANUFACT).ForeColor

395           itmProjSoum.ListSubItems(I_COL_SOUM_MANUFACT).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_MANUFACT).Bold
                   
400           itmProjSoum.SubItems(I_COL_SOUM_PRIX_LIST) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_PRIX_LIST)
405           itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_PRIX_LIST).Tag

410           itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor

415           itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_PRIX_LIST).Bold
          
420           itmProjSoum.SubItems(I_COL_SOUM_ESCOMPTE) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_ESCOMPTE)

425           itmProjSoum.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor

430           itmProjSoum.ListSubItems(I_COL_SOUM_ESCOMPTE).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_ESCOMPTE).Bold
        
435           itmProjSoum.SubItems(I_COL_SOUM_PRIX_NET) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_PRIX_NET)

440           itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor

445           itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_NET).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_PRIX_LIST).Bold
          
450           itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_NET).Tag = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_PRIX_NET).Tag
          
455           itmProjSoum.SubItems(I_COL_SOUM_TOTAL) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_TOTAL)

460           itmProjSoum.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_TOTAL).ForeColor

465           itmProjSoum.ListSubItems(I_COL_SOUM_TOTAL).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_TOTAL).Bold
          
470           itmProjSoum.SubItems(I_COL_SOUM_PROFIT) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_PROFIT)

475           itmProjSoum.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_PROFIT).ForeColor

480           itmProjSoum.ListSubItems(I_COL_SOUM_PROFIT).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_PROFIT).Bold

485           itmProjSoum.ListSubItems(I_COL_SOUM_PROFIT).Tag = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_PROFIT).Tag
            
490           itmProjSoum.SubItems(I_COL_SOUM_DISTRIB) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_DISTRIB)
495           itmProjSoum.ListSubItems(I_COL_SOUM_DISTRIB).Tag = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_DISTRIB).Tag

500           itmProjSoum.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_DISTRIB).ForeColor
          
505           itmProjSoum.ListSubItems(I_COL_SOUM_DISTRIB).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_DISTRIB).Bold
          
510           If Trim$(lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_COMMENTAIRE)) <> "" Then

515             itmProjSoum.SubItems(I_COL_SOUM_COMMENTAIRE) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_COMMENTAIRE)
520             itmProjSoum.ListSubItems(I_COL_SOUM_COMMENTAIRE).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_COMMENTAIRE).ForeColor

525             itmProjSoum.ListSubItems(I_COL_SOUM_COMMENTAIRE).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_COMMENTAIRE).Bold
530           Else
535             itmProjSoum.SubItems(I_COL_SOUM_COMMENTAIRE) = ""
540           End If
          
545           If m_eType = TYPE_PROJET Then
550             If itmProjSoum.SubItems(I_COL_SOUM_PIECE) <> "" And itmProjSoum.SubItems(I_COL_SOUM_PIECE) <> "Texte" And itmProjSoum.SubItems(I_COL_SOUM_PIECE) <> "Text" Then
555               itmProjSoum.SubItems(I_COL_SOUM_FACTURATION) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_FACTURATION)

560               If lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_FACTURATION) = "" Then
565                 lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_FACTURATION) = " "
570               End If

575               itmProjSoum.SubItems(I_COL_SOUM_FACTURATION) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_FACTURATION)

580               itmProjSoum.ListSubItems(I_COL_SOUM_FACTURATION).Tag = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_FACTURATION).Tag

585               itmProjSoum.SubItems(I_COL_SOUM_DATE_COMMANDE) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_DATE_COMMANDE)

590               If itmProjSoum.SubItems(I_COL_SOUM_DATE_COMMANDE) <> "" Then
595                 itmProjSoum.ListSubItems(I_COL_SOUM_DATE_COMMANDE).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_DATE_COMMANDE).ForeColor
600                 itmProjSoum.ListSubItems(I_COL_SOUM_DATE_COMMANDE).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_DATE_COMMANDE).Bold
605                 itmProjSoum.ListSubItems(I_COL_SOUM_DATE_COMMANDE).Tag = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_DATE_COMMANDE).Tag
610               End If

615               itmProjSoum.SubItems(I_COL_SOUM_DATE_REQUISE) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_DATE_REQUISE)

620               If itmProjSoum.SubItems(I_COL_SOUM_DATE_REQUISE) <> "" Then
625                 itmProjSoum.ListSubItems(I_COL_SOUM_DATE_REQUISE).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_DATE_REQUISE).ForeColor
630                 itmProjSoum.ListSubItems(I_COL_SOUM_DATE_REQUISE).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_DATE_REQUISE).Bold
635               End If

640               itmProjSoum.SubItems(I_COL_SOUM_NOM_COMMANDE) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_NOM_COMMANDE)

645               itmProjSoum.ListSubItems(I_COL_SOUM_NOM_COMMANDE).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_NOM_COMMANDE).ForeColor

650               itmProjSoum.ListSubItems(I_COL_SOUM_NOM_COMMANDE).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_NOM_COMMANDE).Bold

655               itmProjSoum.SubItems(I_COL_SOUM_NO_SEQUENTIEL) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_NO_SEQUENTIEL)

660               itmProjSoum.ListSubItems(I_COL_SOUM_NO_SEQUENTIEL).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_NO_SEQUENTIEL).ForeColor

665               itmProjSoum.ListSubItems(I_COL_SOUM_NO_SEQUENTIEL).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_NO_SEQUENTIEL).Bold

670               itmProjSoum.SubItems(I_COL_SOUM_PROVENANCE) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_PROVENANCE)

675               itmProjSoum.ListSubItems(I_COL_SOUM_PROVENANCE).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_PROVENANCE).ForeColor

680               itmProjSoum.ListSubItems(I_COL_SOUM_PROVENANCE).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_PROVENANCE).Bold
685             End If
690           Else
695             If itmProjSoum.SubItems(I_COL_SOUM_PIECE) <> "" And itmProjSoum.SubItems(I_COL_SOUM_PIECE) <> "Texte" And itmProjSoum.SubItems(I_COL_SOUM_PIECE) <> "Text" Then
700               itmProjSoum.SubItems(I_COL_SOUMISSION_PROV) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUMISSION_PROV)
                  
705               If itmProjSoum.SubItems(I_COL_SOUMISSION_PROV) = vbNullString Then
710                 itmProjSoum.SubItems(I_COL_SOUMISSION_PROV) = " "
715               End If
                  
720               itmProjSoum.ListSubItems(I_COL_SOUMISSION_PROV).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUMISSION_PROV).ForeColor

725               itmProjSoum.ListSubItems(I_COL_SOUMISSION_PROV).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUMISSION_PROV).Bold
730             End If
735           End If

740           Call lvwSoumission.ListItems.Remove(iIndexCopie)
                     
745           iIndex = iIndex + 1
750         End If

755         iCompteur2 = iCompteur2 + 1
760       Loop
765     Next

770     If lvwSoumission.ListItems.count > 0 Then
775       Call Deselect

780       lvwSoumission.ListItems(1).Selected = True
785     End If

790     Exit Sub

AfficherErreur:

795     woups "frmProjSoumMec", "UpdateOrdre", Err, Erl
End Sub

Private Function BackupPieces(ByVal sNoProjSoum As String) As Boolean

5       On Error GoTo AfficherErreur

10      Dim rstProjSoum       As ADODB.Recordset
15      Dim rstProjSoumBackup As ADODB.Recordset
20      Dim sDateCopie        As String

25      Set rstProjSoum = New ADODB.Recordset
30      Set rstProjSoumBackup = New ADODB.Recordset

35      If m_eType = TYPE_PROJET Then
40        Call rstProjSoum.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoProjSoum & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

45        Call rstProjSoumBackup.Open("SELECT * FROM GRB_Projet_Pieces_Tampon", g_connData, adOpenDynamic, adLockOptimistic)
50      Else
55        Call rstProjSoum.Open("SELECT * FROM GRB_Soumission_Pieces WHERE IDSoumission = '" & sNoProjSoum & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

60        Call rstProjSoumBackup.Open("SELECT * FROM GRB_Soumission_Pieces_Tampon", g_connData, adOpenDynamic, adLockOptimistic)
65      End If

70      sDateCopie = ConvertDate(Date) & " " & Time

75      Do While Not rstProjSoum.EOF
80        Call rstProjSoumBackup.AddNew

85        rstProjSoumBackup.Fields("DateCopie") = sDateCopie

90        If m_eType = TYPE_PROJET Then
95          rstProjSoumBackup.Fields("IDProjet") = rstProjSoum.Fields("IDProjet")
100       Else
105         rstProjSoumBackup.Fields("IDSoumission") = rstProjSoum.Fields("IDSoumission")
110       End If

115       rstProjSoumBackup.Fields("Initiales") = g_sInitiale
120       rstProjSoumBackup.Fields("IDSection") = rstProjSoum.Fields("IDSection")
125       rstProjSoumBackup.Fields("NumItem") = rstProjSoum.Fields("NumItem")
130       rstProjSoumBackup.Fields("Qté") = rstProjSoum.Fields("Qté")
135       rstProjSoumBackup.Fields("Desc_FR") = rstProjSoum.Fields("Desc_FR")
140       rstProjSoumBackup.Fields("Desc_EN") = rstProjSoum.Fields("Desc_EN")
145       rstProjSoumBackup.Fields("Manufact") = rstProjSoum.Fields("Manufact")
150       rstProjSoumBackup.Fields("Prix_list") = rstProjSoum.Fields("Prix_list")
155       rstProjSoumBackup.Fields("Escompte") = rstProjSoum.Fields("Escompte")
160       rstProjSoumBackup.Fields("Prix_net") = rstProjSoum.Fields("Prix_net")
165       rstProjSoumBackup.Fields("IDFRS") = rstProjSoum.Fields("IDFRS")
170       rstProjSoumBackup.Fields("Temps") = rstProjSoum.Fields("Temps")
175       rstProjSoumBackup.Fields("Temps_total") = rstProjSoum.Fields("Temps_total")
180       rstProjSoumBackup.Fields("Prix_total") = rstProjSoum.Fields("Prix_total")
185       rstProjSoumBackup.Fields("Profit_Argent") = rstProjSoum.Fields("Profit_Argent")
190       rstProjSoumBackup.Fields("SousSection") = rstProjSoum.Fields("sousSection")
195       rstProjSoumBackup.Fields("OrdreSection") = rstProjSoum.Fields("OrdreSection")
200       rstProjSoumBackup.Fields("NuméroLigne") = rstProjSoum.Fields("NuméroLigne")
205       rstProjSoumBackup.Fields("PrixOrigine") = rstProjSoum.Fields("PrixOrigine")
210       rstProjSoumBackup.Fields("Type") = rstProjSoum.Fields("Type")
215       rstProjSoumBackup.Fields("Visible") = rstProjSoum.Fields("Visible")
220       rstProjSoumBackup.Fields("Commandé") = rstProjSoum.Fields("Commandé")
225       rstProjSoumBackup.Fields("Quoté") = rstProjSoum.Fields("Quoté")
230       rstProjSoumBackup.Fields("Recu") = rstProjSoum.Fields("Recu")
235       rstProjSoumBackup.Fields("Retour") = rstProjSoum.Fields("Retour")
240       rstProjSoumBackup.Fields("CommandeAnnulée") = rstProjSoum.Fields("CommandeAnnulée")
245       rstProjSoumBackup.Fields("ID") = rstProjSoum.Fields("ID")
250       rstProjSoumBackup.Fields("PieceExtra") = rstProjSoum.Fields("PieceExtra")
255       rstProjSoumBackup.Fields("PieceExtraChargeable") = rstProjSoum.Fields("PieceExtraChargeable")
260       rstProjSoumBackup.Fields("PieceExtraNonChargeable") = rstProjSoum.Fields("PieceExtraNonChargeable")
265       rstProjSoumBackup.Fields("MatérielInutile") = rstProjSoum.Fields("MatérielInutile")
270       rstProjSoumBackup.Fields("Commentaire") = rstProjSoum.Fields("Commentaire")
275       rstProjSoumBackup.Fields("Devise") = rstProjSoum.Fields("Devise")

280       If m_eType = TYPE_PROJET Then
285         rstProjSoumBackup.Fields("NoRetour") = rstProjSoum.Fields("NoRetour")
290         rstProjSoumBackup.Fields("DateRéception") = rstProjSoum.Fields("DateRéception")
295         rstProjSoumBackup.Fields("QuantitéRecue") = rstProjSoum.Fields("QuantitéRecue")
300         rstProjSoumBackup.Fields("Facturation") = rstProjSoum.Fields("Facturation")
305         rstProjSoumBackup.Fields("DateCommande") = rstProjSoum.Fields("DateCommande")
310         rstProjSoumBackup.Fields("DateRequise") = rstProjSoum.Fields("DateRequise")
315         rstProjSoumBackup.Fields("NomCommande") = rstProjSoum.Fields("NomCommande")
320         rstProjSoumBackup.Fields("NoSéquentiel") = rstProjSoum.Fields("NoSéquentiel")
325         rstProjSoumBackup.Fields("DateRetour") = rstProjSoum.Fields("DateRetour")
330       End If

335       rstProjSoumBackup.Fields("Provenance") = rstProjSoum.Fields("Provenance")

340       Call rstProjSoumBackup.Update

345       Call rstProjSoum.MoveNext
350     Loop

355     Call rstProjSoum.Close
360     Set rstProjSoum = Nothing

365     Call rstProjSoumBackup.Close
370     Set rstProjSoumBackup = Nothing

375     BackupPieces = True

380     Exit Function

AfficherErreur:

385     woups "frmProjSoumMec", "BackupPieces", Err, Erl
End Function

Private Sub UpdatePieces()

5       On Error GoTo AfficherErreur

10      Dim rstPieceFRS As ADODB.Recordset
15      Dim rstConfig   As ADODB.Recordset
20      Dim itmPiece    As ListItem
25      Dim iCompteur   As Integer
30      Dim sTauxUSA    As String
35      Dim sTauxSPA    As String
  
40      Set rstPieceFRS = New ADODB.Recordset
  
45      Set rstConfig = New ADODB.Recordset

50      Call rstConfig.Open("SELECT TauxAmericain, TauxEspagnol FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)
  
55      sTauxUSA = rstConfig.Fields("TauxAmericain")
60      sTauxSPA = rstConfig.Fields("TauxEspagnol")

65      Call rstConfig.Close
70      Set rstConfig = Nothing
  
75      For iCompteur = 1 To lvwSoumission.ListItems.count
          'Si ce n'est pas une section
80        If lvwSoumission.ListItems(iCompteur).Tag <> vbNullString Then
            'Si ce n'est pas une sous-section
85          If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> vbNullString Then
90            Set itmPiece = lvwSoumission.ListItems(iCompteur)
                
95            If itmPiece.SubItems(I_COL_SOUM_DISTRIB) <> vbNullString Then
100             Call ValeurParDefaut(itmPiece)

105             Call rstPieceFRS.Open("SELECT PRIX_LIST, PRIX_SP, PRIX_NET, ESCOMPTE, DeviseMonétaire FROM GRB_PiecesFRS WHERE PIECE = '" & Replace(itmPiece.SubItems(I_COL_SOUM_PIECE), "'", "''") & "' AND IDFRS = " & itmPiece.ListSubItems(I_COL_SOUM_DISTRIB).Tag, g_connData, adOpenDynamic, adLockOptimistic)

110             If Not rstPieceFRS.EOF Then
115               If Not IsNull(rstPieceFRS.Fields("PRIX_LIST")) Then
120                 If Trim(rstPieceFRS.Fields("PRIX_LIST")) <> "" Then
125                   If rstPieceFRS.Fields("DeviseMonétaire") = "USA" Then
130                     itmPiece.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(rstPieceFRS.Fields("PRIX_LIST")) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
135                   Else
140                     If rstPieceFRS.Fields("DeviseMonétaire") = "SPA" Then
145                       itmPiece.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(rstPieceFRS.Fields("PRIX_LIST")) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
150                     Else
155                       itmPiece.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(rstPieceFRS.Fields("PRIX_LIST"), MODE_ARGENT, 4)
160                     End If
165                   End If
170                 Else
175                   itmPiece.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion("0", MODE_ARGENT, 4)
180                 End If
185               Else
190                 itmPiece.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion("0", MODE_ARGENT, 4)
195               End If
            
200               If Trim(rstPieceFRS.Fields("PRIX_NET")) <> vbNullString Then
205                 If rstPieceFRS.Fields("ESCOMPTE") <> vbNullString Then
210                   itmPiece.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion(rstPieceFRS.Fields("Escompte"), MODE_POURCENT)
215                 End If
           
220                 itmPiece.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(rstPieceFRS.Fields("PRIX_NET"), MODE_ARGENT, 4)
225               Else
230                 itmPiece.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(rstPieceFRS.Fields("PRIX_SP"), MODE_ARGENT, 4)
235               End If

240               If rstPieceFRS.Fields("DeviseMonétaire") = "USA" Then
245                 itmPiece.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(itmPiece.SubItems(I_COL_SOUM_PRIX_NET)) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
250               Else
255                 If rstPieceFRS.Fields("DeviseMonétaire") = "SPA" Then
260                   itmPiece.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(itmPiece.SubItems(I_COL_SOUM_PRIX_NET)) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
265                 Else
270                   itmPiece.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(itmPiece.SubItems(I_COL_SOUM_PRIX_NET), MODE_ARGENT, 4)
275                 End If
280               End If
      
                  'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
285               itmPiece.SubItems(I_COL_SOUM_TOTAL) = Conversion(CStr(Round(Replace(itmPiece.Text, "*", "") * itmPiece.SubItems(I_COL_SOUM_PRIX_NET) * CSng(m_sProfit), 2)), MODE_ARGENT)
     
                  'Pour le profit, c'est le prix total - (prix net * quantité)
290               itmPiece.SubItems(I_COL_SOUM_PROFIT) = Conversion(CStr(Round(itmPiece.SubItems(I_COL_SOUM_TOTAL) - (itmPiece.SubItems(I_COL_SOUM_PRIX_NET) * Replace(itmPiece.Text, "*", "")), 2)), MODE_ARGENT)
     
                  'Pour garder en mémoire le prix d'origine, je le mets dans le
                  'tag de la colonne Prix Listé
295               If itmPiece.SubItems(I_COL_SOUM_PRIX_LIST) = vbNullString Then
300                 itmPiece.SubItems(I_COL_SOUM_PRIX_LIST) = " "
305               End If
     
310               If rstPieceFRS.Fields("PRIX_NET") <> vbNullString Then
315                 itmPiece.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = rstPieceFRS.Fields("PRIX_LIST")
320               Else
325                 itmPiece.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = rstPieceFRS.Fields("PRIX_SP")
330               End If
335             Else
340               Call MsgBox("Il n'y a pas de prix du fournisseur " & itmPiece.SubItems(I_COL_SOUM_DISTRIB) & " pour la pièce " & itmPiece.SubItems(I_COL_SOUM_PIECE) & " ou la pièce n'existe plus!", vbOKOnly, "Erreur")
345             End If

350             Call rstPieceFRS.Close
355           End If
360         End If
365       End If
370     Next

375     Set rstPieceFRS = Nothing

380     Exit Sub

AfficherErreur:

385     woups "frmProjSoumMec", "UpdatePieces", Err, Erl
End Sub

Private Sub cmdCreerProjet_Click()

5       On Error GoTo AfficherErreur

        'Créé un projet à partir d'une soumission
10      Dim rstProjSoum As ADODB.Recordset
15      Dim sNoProjet   As String
20      Dim sUser       As String
25      Dim iCompteur   As Integer
30      Dim bExiste     As Boolean
35      Dim bNoValide   As Boolean
40      Dim sLiaison    As String

45      If cmbProjSoum.ListCount > 0 Then
50        If Right$(txtNoProjSoum.Text, 2) = "99" Then
55          Call MsgBox("Impossible de créer un projet à partir de cette soumission!", vbOKOnly, "Erreur")
        
60          Exit Sub
65        End If

70        Set rstProjSoum = New ADODB.Recordset

75        Call rstProjSoum.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

80        If rstProjSoum.Fields("Ouvert") = False Or rstProjSoum.Fields("Verrouillé") = True Then
85          If rstProjSoum.Fields("Ouvert") = False Then
90            Call MsgBox("Cette soumission est fermée!", vbOKOnly)
95          Else
100           Call MsgBox("Cette soumission est verrouillée!", vbOKOnly)
105         End If

110         Call rstProjSoum.Close
115         Set rstProjSoum = Nothing

120         Exit Sub
125       End If

130       Call rstProjSoum.Close

135       If VerifierSiOuvert(sUser) = False Then
            'Demande du numéro de projet
140         sNoProjet = InputBox("Quel est le numéro du projet?")
  
145         If sNoProjet <> vbNullString Then
150           Screen.MousePointer = vbHourglass

155           bNoValide = True

160           If ValiderFormatNumeroProjSoum(sNoProjet) = False Then
165             bNoValide = False
170           End If

175           If bNoValide = True Then
180             If ValiderFormatMecanique(sNoProjet) = False Then
185               bNoValide = False
190             End If
195           End If

200           If bNoValide = True Then
205             If ValiderFormatJobAvecSoum(sNoProjet) = False Then
210               bNoValide = False
215             End If
220           End If

225           If bNoValide = False Then
230             Set rstProjSoum = Nothing

235             Screen.MousePointer = vbDefault

240             Exit Sub
245           End If

250           sNoProjet = UCase(sNoProjet)
  
255           Call rstProjSoum.Open("SELECT * FROM GRB_SoumissionElec WHERE IDSoumission = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)
        
260           If rstProjSoum.EOF Then
265             bExiste = False
270           Else
275             bExiste = True
  
280             Call MsgBox("Ce numéro existe dans les soumissions électriques!", vbOKOnly, "Erreur")
285           End If

290           Call rstProjSoum.Close

295           If bExiste = False Then
300             Call rstProjSoum.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)
            
305             If rstProjSoum.EOF Then
310               bExiste = False
315             Else
320               bExiste = True
  
325               Call MsgBox("Ce numéro existe dans les projets électriques!", vbOKOnly, "Erreur")
330             End If

335             Call rstProjSoum.Close
340           End If
  
345           If bExiste = False Then
350             Call rstProjSoum.Open("SELECT * FROM GRB_SoumissionMec WHERE IDSoumission = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)
  
355             If rstProjSoum.EOF Then
360               bExiste = False
365             Else
370               bExiste = True
    
375               Call MsgBox("Ce numéro existe dans les soumissions mécaniques!", vbOKOnly, "Erreur")
380             End If

385             Call rstProjSoum.Close
390           End If
         
395           If bExiste = False Then
400             Call rstProjSoum.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)
    
405             If rstProjSoum.EOF Then
410               bExiste = False
415             Else
420               bExiste = True
  
425               Call MsgBox("Ce numéro existe dans les projets mécaniques!", vbOKOnly, "Erreur")
430             End If
  
435             Call rstProjSoum.Close
440           End If

445           If bExiste = True Then
450             Set rstProjSoum = Nothing

455             Screen.MousePointer = vbDefault

460             Exit Sub
465           End If

470           Call rstProjSoum.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)

475           If Not rstProjSoum.EOF Then
480             If rstProjSoum.Fields("Ouvert") = False Then
485               Call MsgBox("Ce numéro est fermé!", vbOKOnly, "Erreur")
                   
490               Call rstProjSoum.Close
495               Set rstProjSoum = Nothing

500               Exit Sub
505             End If
510           End If

515           Call rstProjSoum.Close
520           Set rstProjSoum = Nothing
      
525           If Right$(sNoProjet, 2) >= 60 And Right$(sNoProjet, 2) <= 98 Then
530             sLiaison = InputBox("Quelle est l'extention du projet " & Left$(sNoProjet, 6) & " auquel ce projet sera lié?")
535           End If
      
540           Call frmChoixTransfertJob.Afficher(txtNoProjSoum.Text, "M")

545           If m_bTransfertJobCancel = False Then
                'Appel de la méthode pour créer le projet
550             Call TransfererSoumDansProjet(sNoProjet, sLiaison)
      
                'On affiche le projet qui vient d'être créé
555             If m_bComboChoix = True Then
560               cmbChoix.ListIndex = I_IDX_PROJET
             
565               For iCompteur = 0 To cmbProjSoum.ListCount - 1
570                 If cmbProjSoum.LIST(iCompteur) = sNoProjet Then
575                   cmbProjSoum.ListIndex = iCompteur
                    
580                   Exit For
585                 End If
590               Next

595               If sLiaison <> "" Then
600                 For iCompteur = 1 To lvwSoumission.ListItems.count
                      'Si pas une section
605                   If lvwSoumission.ListItems(iCompteur).Tag <> "" Then
                        'Si pas une sous-section
610                     If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> "" Then
615                       If Right$(sNoProjet, 2) >= 60 And Right$(sNoProjet, 2) <= 79 Then
620                         Call AjouterPiecesExtraChargeableDansJob(lvwSoumission.ListItems(iCompteur), sLiaison)
625                       Else
630                         If Right$(sNoProjet, 2) >= 80 And Right$(sNoProjet, 2) <= 98 Then
635                           Call AjouterPiecesExtraDansJob(lvwSoumission.ListItems(iCompteur), sLiaison)
640                         End If
645                       End If
650                     End If
655                   End If

660                   Call CalculerTotalRecordset(sNoProjet)
665                 Next
670               End If

675               Call AjouterProjetAuCumulatif
680             End If
        
                'Il faut enlever le bouton puisqu'on ne peut pas créer plus
                'd'un projet avec une soumission
685             cmdCreerProjet.Visible = False
690           End If
    
695           Screen.MousePointer = vbDefault
700         Else
705           Set rstProjSoum = Nothing

710           Exit Sub
715         End If
720       Else
725         Call MsgBox("Cette soumission est en modification par " & sUser & "!", vbOKOnly, "Erreur")
730       End If
735     End If

740     Exit Sub

AfficherErreur:

745     woups "frmProjSoumMec", "cmdCreerProjet_Click", Err, Erl
End Sub

Private Function VerifierSiDejaProjet() As Boolean

5       On Error GoTo AfficherErreur

        'Méthode qui sert à vérifier si une soumission est déjà assignée à un projet
10      Dim rstProjet As ADODB.Recordset
  
15      If txtNoProjSoum.Text = vbNullString Then
20        VerifierSiDejaProjet = True
    
25        Exit Function
30      End If
  
35      Set rstProjet = New ADODB.Recordset
  
40      Call rstProjet.Open("SELECT * FROM GRB_ProjetMec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
    
45      If Not rstProjet.EOF Then
50        VerifierSiDejaProjet = True
55      End If
    
60      Call rstProjet.Close
65      Set rstProjet = Nothing

70      Exit Function

AfficherErreur:

75      woups "frmProjSoumMec", "VerifierSiDejaProjet", Err, Erl
End Function

Private Sub TransfererSoumDansProjet(ByVal sNoProjet As String, ByVal sLiaison As String)

5       On Error GoTo AfficherErreur

        'Méthode qui transfère les données de la soumission dans les tables
        'GRB_Projet et GRB_pièces
10      Dim rstSoum        As ADODB.Recordset
15      Dim rstProjet      As ADODB.Recordset
20      Dim rstSoumPiece   As ADODB.Recordset
25      Dim rstProjetPiece As ADODB.Recordset
30      Dim rstEmploye     As ADODB.Recordset
35      Dim rstProjSoum    As ADODB.Recordset
40      Dim rstConfig      As ADODB.Recordset
45      Dim iCompteur      As Integer

50      Set rstSoum = New ADODB.Recordset
55      Set rstSoumPiece = New ADODB.Recordset

        'Ouverture de la soumission
60      Call rstSoum.Open("SELECT * FROM GRB_SoumissionMec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
65      Call rstSoumPiece.Open("SELECT * FROM GRB_Soumission_Pieces WHERE IDSoumission = '" & txtNoProjSoum.Text & "' AND Type = 'M' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
    
70      Set rstProjet = New ADODB.Recordset
75      Set rstProjetPiece = New ADODB.Recordset
    
        'Ouverture du projet
80      Call rstProjet.Open("SELECT * FROM GRB_ProjetMec", g_connData, adOpenDynamic, adLockOptimistic)
85      Call rstProjetPiece.Open("SELECT * FROM GRB_Projet_Pieces", g_connData, adOpenDynamic, adLockOptimistic)
    
90      Set rstProjSoum = New ADODB.Recordset
    
95      Call rstProjSoum.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)

100     If rstProjSoum.EOF Then
105       Call rstProjSoum.AddNew
    
110       rstProjSoum.Fields("IDProjSoum") = sNoProjet
115       rstProjSoum.Fields("NoClient") = rstSoum.Fields("IDClient")
120       rstProjSoum.Fields("Description") = rstSoum.Fields("Description")
125       rstProjSoum.Fields("DateOuverture") = ConvertDate(Date)
130       rstProjSoum.Fields("Ouvert") = True
135       rstProjSoum.Fields("Type") = "P"
    
140       Call rstProjSoum.Update
145     End If
    
150     Call rstProjSoum.Close

155     Call rstProjSoum.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

160     rstProjSoum.Fields("Ouvert") = False

165     Call rstProjSoum.Update

170     Call rstProjSoum.Close
175     Set rstProjSoum = Nothing
  
        'On l'ajoute
180     Call rstProjet.AddNew
       
185     rstProjet.Fields("IDProjet") = sNoProjet
190     rstProjet.Fields("IDSoumission") = rstSoum.Fields("IDSoumission")
195     rstProjet.Fields("IDClient") = rstSoum.Fields("IDClient")
200     rstProjet.Fields("IDContact") = rstSoum.Fields("IDContact")
205     rstProjet.Fields("Description") = rstSoum.Fields("Description")
210     rstProjet.Fields("manuel") = rstSoum.Fields("manuel")
215     rstProjet.Fields("Creer") = ConvertDate(Date)
220     rstProjet.Fields("CheminPhotos") = rstSoum.Fields("CheminPhotos")
     
225     If sLiaison <> "" Then
230       rstProjet.Fields("LiaisonChargeable") = sLiaison
235     End If
     
240     Set rstEmploye = New ADODB.Recordset
     
245     Call rstEmploye.Open("SELECT NoEmploye FROM GRB_Employés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)
    
250     rstProjet.Fields("Creer_Par") = rstEmploye.Fields("noEmploye")
     
255     Call rstEmploye.Close
260     Set rstEmploye = Nothing
      
265     rstProjet.Fields("imprevue") = rstSoum.Fields("imprevue")
270     rstProjet.Fields("commission") = rstSoum.Fields("commission")
275     rstProjet.Fields("Profit") = rstSoum.Fields("Profit")
280     rstProjet.Fields("total_manuel") = rstSoum.Fields("total_manuel")
285     rstProjet.Fields("total_commission") = rstSoum.Fields("total_Commission")
290     rstProjet.Fields("total_profit") = rstSoum.Fields("Total_Profit")
295     rstProjet.Fields("PrixTotal") = rstSoum.Fields("PrixTotal")
300     rstProjet.Fields("total_piece") = rstSoum.Fields("Total_piece")
305     rstProjet.Fields("total_imprevue") = rstSoum.Fields("total_imprevue")
310     rstProjet.Fields("total_temps") = rstSoum.Fields("total_temps")

315     rstProjet.Fields("TempsDessinProj") = rstSoum.Fields("TempsDessin")
320     rstProjet.Fields("TempsCoupeProj") = rstSoum.Fields("TempsCoupe")
325     rstProjet.Fields("TempsMachinageProj") = rstSoum.Fields("TempsMachinage")
330     rstProjet.Fields("TempsSoudureProj") = rstSoum.Fields("TempsSoudure")
335     rstProjet.Fields("TempsAssemblageProj") = rstSoum.Fields("TempsAssemblage")
340     rstProjet.Fields("TempsPeintureProj") = rstSoum.Fields("TempsPeinture")
345     rstProjet.Fields("TempsTestProj") = rstSoum.Fields("TempsTest")
350     rstProjet.Fields("TempsInstallationProj") = 0
355     rstProjet.Fields("TempsFormationProj") = rstSoum.Fields("TempsFormation")
360     rstProjet.Fields("TempsGestionProj") = rstSoum.Fields("TempsGestion")
365     rstProjet.Fields("TempsShippingProj") = rstSoum.Fields("TempsShipping")

370     rstProjet.Fields("TempsDessinConc") = vbNullString
375     rstProjet.Fields("TempsCoupeConc") = vbNullString
380     rstProjet.Fields("TempsMachinageConc") = vbNullString
385     rstProjet.Fields("TempsSoudureConc") = vbNullString
390     rstProjet.Fields("TempsAssemblageConc") = vbNullString
395     rstProjet.Fields("TempsPeintureConc") = vbNullString
400     rstProjet.Fields("TempsTestConc") = vbNullString
405     rstProjet.Fields("TempsInstallationConc") = vbNullString
410     rstProjet.Fields("TempsFormationConc") = vbNullString
415     rstProjet.Fields("TempsGestionConc") = vbNullString
420     rstProjet.Fields("TempsShippingConc") = vbNullString

425     rstProjet.Fields("PrixEmballage") = rstSoum.Fields("PrixEmballage")

430     Set rstConfig = New ADODB.Recordset

435     If Not IsNull(rstSoum.Fields("TauxDessin")) Then
440       rstProjet.Fields("TauxDessin") = rstSoum.Fields("TauxDessin")
445     Else
450       Call rstConfig.Open("SELECT TauxDessinMec FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

455       rstProjet.Fields("TauxDessin") = rstConfig.Fields("TauxDessinMec")

460       Call rstConfig.Close
465     End If

470     If Not IsNull(rstSoum.Fields("TauxCoupe")) Then
475       rstProjet.Fields("TauxCoupe") = rstSoum.Fields("TauxCoupe")
480     Else
485       Call rstConfig.Open("SELECT TauxCoupe FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

490       rstProjet.Fields("TauxCoupe") = rstConfig.Fields("TauxCoupe")

495       Call rstConfig.Close
500     End If

505     If Not IsNull(rstSoum.Fields("TauxMachinage")) Then
510       rstProjet.Fields("TauxMachinage") = rstSoum.Fields("TauxMachinage")
515     Else
520       Call rstConfig.Open("SELECT TauxMachinage FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

525       rstProjet.Fields("TauxMachinage") = rstConfig.Fields("TauxMachinage")

530       Call rstConfig.Close
535     End If

540     If Not IsNull(rstSoum.Fields("TauxSoudure")) Then
545       rstProjet.Fields("TauxSoudure") = rstSoum.Fields("TauxSoudure")
550     Else
555       Call rstConfig.Open("SELECT TauxSoudure FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

560       rstProjet.Fields("TauxSoudure") = rstConfig.Fields("TauxSoudure")

565       Call rstConfig.Close
570     End If

575     If Not IsNull(rstSoum.Fields("TauxAssemblage")) Then
580       rstProjet.Fields("TauxAssemblage") = rstSoum.Fields("TauxAssemblage")
585     Else
590       Call rstConfig.Open("SELECT TauxAssemblageMec FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

595       rstProjet.Fields("TauxAssemblage") = rstConfig.Fields("TauxAssemblageMec")

600       Call rstConfig.Close
605     End If

610     If Not IsNull(rstSoum.Fields("TauxPeinture")) Then
615       rstProjet.Fields("TauxPeinture") = rstSoum.Fields("TauxPeinture")
620     Else
625       Call rstConfig.Open("SELECT TauxPeinture FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

630       rstProjet.Fields("TauxPeinture") = rstConfig.Fields("TauxPeinture")

635       Call rstConfig.Close
640     End If

645     If Not IsNull(rstSoum.Fields("TauxTest")) Then
650       rstProjet.Fields("TauxTest") = rstSoum.Fields("TauxTest")
655     Else
660       Call rstConfig.Open("SELECT TauxTestMec FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

665       rstProjet.Fields("TauxTest") = rstConfig.Fields("TauxTestMec")

670       Call rstConfig.Close
675     End If

680     If Not IsNull(rstSoum.Fields("TauxInstallation")) Then
685       rstProjet.Fields("TauxInstallation") = rstSoum.Fields("TauxInstallation")
690     Else
695       Call rstConfig.Open("SELECT TauxInstallationMec FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

700       rstProjet.Fields("TauxInstallation") = rstConfig.Fields("TauxInstallationMec")

705       Call rstConfig.Close
710     End If

715     If Not IsNull(rstSoum.Fields("TauxFormation")) Then
720       rstProjet.Fields("TauxFormation") = rstSoum.Fields("TauxFormation")
725     Else
730       Call rstConfig.Open("SELECT TauxFormationMec FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

735       rstProjet.Fields("TauxFormation") = rstConfig.Fields("TauxFormationMec")

740       Call rstConfig.Close
745     End If

750     If Not IsNull(rstSoum.Fields("TauxGestion")) Then
755       rstProjet.Fields("TauxGestion") = rstSoum.Fields("TauxGestion")
760     Else
765       Call rstConfig.Open("SELECT TauxGestionProjetsMec FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

770       rstProjet.Fields("TauxGestion") = rstConfig.Fields("TauxGestionProjetsMec")

775       Call rstConfig.Close
780     End If

785     If Not IsNull(rstSoum.Fields("TauxShipping")) Then
790       rstProjet.Fields("TauxShipping") = rstSoum.Fields("TauxShipping")
795     Else
800       Call rstConfig.Open("SELECT TauxShippingMec FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

805       rstProjet.Fields("TauxShipping") = rstConfig.Fields("TauxShippingMec")

810       Call rstConfig.Close
815     End If

820     Set rstConfig = Nothing

825     rstProjet.Fields("MontantForfait") = rstSoum.Fields("MontantForfait")
830     rstProjet.Fields("InitialeForfait") = rstSoum.Fields("InitialeForfait")
835     rstProjet.Fields("ProchaineCommande") = 1
   
840     Call rstProjet.Update
    
        'Ajout des pièces
845     Do While Not rstSoumPiece.EOF
850       If rstSoumPiece.Fields("TransfertJob") = True Then
855         Call rstProjetPiece.AddNew
  
860         rstProjetPiece.Fields("IDProjet") = sNoProjet
865         rstProjetPiece.Fields("IDSection") = rstSoumPiece.Fields("IDSection")
870         rstProjetPiece.Fields("NumItem") = rstSoumPiece.Fields("NumItem")
875         rstProjetPiece.Fields("Qté") = rstSoumPiece.Fields("Qté")
880         rstProjetPiece.Fields("Desc_FR") = rstSoumPiece.Fields("Desc_FR")
885         rstProjetPiece.Fields("Desc_EN") = rstSoumPiece.Fields("Desc_EN")
890         rstProjetPiece.Fields("Manufact") = rstSoumPiece.Fields("Manufact")
895         rstProjetPiece.Fields("Prix_List") = rstSoumPiece.Fields("Prix_list")
900         rstProjetPiece.Fields("Escompte") = rstSoumPiece.Fields("Escompte")
905         rstProjetPiece.Fields("Prix_net") = rstSoumPiece.Fields("Prix_net")
910         rstProjetPiece.Fields("OrdreSection") = rstSoumPiece.Fields("OrdreSection")
915         rstProjetPiece.Fields("NuméroLigne") = rstSoumPiece.Fields("NuméroLigne")
920         rstProjetPiece.Fields("IDFRS") = rstSoumPiece.Fields("IDFRS")
925         rstProjetPiece.Fields("Prix_total") = rstSoumPiece.Fields("Prix_Total")
930         rstProjetPiece.Fields("Profit_argent") = rstSoumPiece.Fields("Profit_argent")
935         rstProjetPiece.Fields("SousSection") = rstSoumPiece.Fields("SousSection")
940         rstProjetPiece.Fields("PrixOrigine") = rstSoumPiece.Fields("PrixOrigine")
945         rstProjetPiece.Fields("Visible") = rstSoumPiece.Fields("Visible")
950         rstProjetPiece.Fields("Commentaire") = rstSoumPiece.Fields("Commentaire")
955         rstProjetPiece.Fields("Quoté") = rstSoumPiece.Fields("Quoté")

960         rstProjetPiece.Fields("Type") = rstSoumPiece.Fields("Type")
      
965         Call rstProjetPiece.Update
970       End If
   
975       Call rstSoumPiece.MoveNext
980     Loop

985     m_eType = TYPE_PROJET

990     If CDbl(rstSoum.Fields("TempsInstallation")) > 0 Then
995       Call CreerProjetInstallation(Left$(sNoProjet, 7) & "51")
1000    End If

1005    Call rstSoum.Close
1010    Set rstSoum = Nothing
    
1015    Call rstProjet.Close
1020    Set rstProjet = Nothing
    
1025    Call rstSoumPiece.Close
1030    Set rstSoumPiece = Nothing
   
1035    Call rstProjetPiece.Close
1040    Set rstProjetPiece = Nothing

1045    Call CalculerTotalRecordset(sNoProjet)

1050    Exit Sub

AfficherErreur:

1055    woups "frmProjSoumMec", "TransfererSoumDansProjet", Err, Erl
End Sub

Private Sub CreerProjetInstallation(ByVal sNoProjet As String)
  
5       On Error GoTo AfficherErreur

10      Dim rstSoum     As ADODB.Recordset
15      Dim rstProjet   As ADODB.Recordset
20      Dim rstEmploye  As ADODB.Recordset
25      Dim rstProjSoum As ADODB.Recordset
30      Dim rstConfig   As ADODB.Recordset
35      Dim iCompteur   As Integer

40      Set rstSoum = New ADODB.Recordset

        'Ouverture de la soumission
45      Call rstSoum.Open("SELECT * FROM GRB_SoumissionMec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
    
50      Set rstProjet = New ADODB.Recordset
        
        'Ouverture du projet
55      Call rstProjet.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)
    
60      If rstProjet.EOF Then
65        Set rstProjSoum = New ADODB.Recordset
    
70        Call rstProjSoum.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)

75        If rstProjSoum.EOF Then
80          Call rstProjSoum.AddNew
    
85          rstProjSoum.Fields("IDProjSoum") = sNoProjet
90          rstProjSoum.Fields("NoClient") = rstSoum.Fields("IDClient")
95          rstProjSoum.Fields("DateOuverture") = ConvertDate(Date)
100         rstProjSoum.Fields("Description") = rstSoum.Fields("Description")
105         rstProjSoum.Fields("Ouvert") = True
110         rstProjSoum.Fields("Type") = "P"
    
115         Call rstProjSoum.Update
120       End If
    
125       Call rstProjSoum.Close
130       Set rstProjSoum = Nothing
  
          'On l'ajoute
135       Call rstProjet.AddNew
       
140       rstProjet.Fields("IDProjet") = sNoProjet
145       rstProjet.Fields("IDSoumission") = vbNullString
150       rstProjet.Fields("IDClient") = rstSoum.Fields("IDClient")
155       rstProjet.Fields("IDContact") = rstSoum.Fields("IDContact")
160       rstProjet.Fields("Description") = rstSoum.Fields("Description")
165       rstProjet.Fields("Creer") = ConvertDate(Date)
170       rstProjet.Fields("CheminPhotos") = rstSoum.Fields("CheminPhotos")
     
175       Set rstEmploye = New ADODB.Recordset
     
180       Call rstEmploye.Open("SELECT NoEmploye FROM GRB_Employés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)
    
185       rstProjet.Fields("Creer_Par") = rstEmploye.Fields("noEmploye")
     
190       Call rstEmploye.Close
195       Set rstEmploye = Nothing
      
200       rstProjet.Fields("imprevue") = rstSoum.Fields("imprevue")
205       rstProjet.Fields("commission") = rstSoum.Fields("commission")
210       rstProjet.Fields("Profit") = rstSoum.Fields("Profit")
215       rstProjet.Fields("total_manuel") = rstSoum.Fields("total_manuel")
220       rstProjet.Fields("total_commission") = rstSoum.Fields("total_Commission")
225       rstProjet.Fields("total_profit") = rstSoum.Fields("Total_Profit")
230       rstProjet.Fields("PrixTotal") = rstSoum.Fields("PrixTotal")
235       rstProjet.Fields("total_piece") = rstSoum.Fields("Total_piece")
240       rstProjet.Fields("total_imprevue") = rstSoum.Fields("total_imprevue")
245       rstProjet.Fields("total_temps") = rstSoum.Fields("total_temps")

250       rstProjet.Fields("TempsDessinProj") = 0
255       rstProjet.Fields("TempsCoupeProj") = 0
260       rstProjet.Fields("TempsMachinageProj") = 0
265       rstProjet.Fields("TempsSoudureProj") = 0
270       rstProjet.Fields("TempsAssemblageProj") = 0
275       rstProjet.Fields("TempsPeintureProj") = 0
280       rstProjet.Fields("TempsTestProj") = 0
285       rstProjet.Fields("TempsInstallationProj") = rstSoum.Fields("TempsInstallation")
290       rstProjet.Fields("TempsFormationProj") = 0
295       rstProjet.Fields("TempsGestionProj") = 0
300       rstProjet.Fields("TempsShippingProj") = 0

305       rstProjet.Fields("TempsDessinConc") = vbNullString
310       rstProjet.Fields("TempsCoupeConc") = vbNullString
315       rstProjet.Fields("TempsMachinageConc") = vbNullString
320       rstProjet.Fields("TempsSoudureConc") = vbNullString
325       rstProjet.Fields("TempsAssemblageConc") = vbNullString
330       rstProjet.Fields("TempsPeintureConc") = vbNullString
335       rstProjet.Fields("TempsTestConc") = vbNullString
340       rstProjet.Fields("TempsInstallationConc") = vbNullString
345       rstProjet.Fields("TempsFormationConc") = vbNullString
350       rstProjet.Fields("TempsGestionConc") = vbNullString
355       rstProjet.Fields("TempsShippingConc") = vbNullString

360       rstProjet.Fields("PrixEmballage") = 0

365       Set rstConfig = New ADODB.Recordset

370       If Not IsNull(rstSoum.Fields("TauxDessin")) Then
375         rstProjet.Fields("TauxDessin") = rstSoum.Fields("TauxDessin")
380       Else
385         Call rstConfig.Open("SELECT TauxDessinMec FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

390         rstProjet.Fields("TauxDessin") = rstConfig.Fields("TauxDessinMec")
  
395         Call rstConfig.Close
400       End If

405       If Not IsNull(rstSoum.Fields("TauxCoupe")) Then
410         rstProjet.Fields("TauxCoupe") = rstSoum.Fields("TauxCoupe")
415       Else
420         Call rstConfig.Open("SELECT TauxCoupe FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

425         rstProjet.Fields("TauxCoupe") = rstConfig.Fields("TauxCoupe")
  
430         Call rstConfig.Close
435       End If

440       If Not IsNull(rstSoum.Fields("TauxMachinage")) Then
445         rstProjet.Fields("TauxMachinage") = rstSoum.Fields("TauxMachinage")
450       Else
455         Call rstConfig.Open("SELECT TauxMachinage FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

460         rstProjet.Fields("TauxMachinage") = rstConfig.Fields("TauxMachinage")
  
465         Call rstConfig.Close
470       End If

475       If Not IsNull(rstSoum.Fields("TauxSoudure")) Then
480         rstProjet.Fields("TauxSoudure") = rstSoum.Fields("TauxSoudure")
485       Else
490         Call rstConfig.Open("SELECT TauxSoudure FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)
  
495         rstProjet.Fields("TauxSoudure") = rstConfig.Fields("TauxSoudure")

500         Call rstConfig.Close
505       End If
 
510       If Not IsNull(rstSoum.Fields("TauxAssemblage")) Then
515         rstProjet.Fields("TauxAssemblage") = rstSoum.Fields("TauxAssemblage")
520       Else
525         Call rstConfig.Open("SELECT TauxAssemblageMec FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

530         rstProjet.Fields("TauxAssemblage") = rstConfig.Fields("TauxAssemblageMec")

535         Call rstConfig.Close
540       End If

545       If Not IsNull(rstSoum.Fields("TauxPeinture")) Then
550         rstProjet.Fields("TauxPeinture") = rstSoum.Fields("TauxPeinture")
555       Else
560         Call rstConfig.Open("SELECT TauxPeinture FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

565         rstProjet.Fields("TauxPeinture") = rstConfig.Fields("TauxPeinture")

570         Call rstConfig.Close
575       End If

580       If Not IsNull(rstSoum.Fields("TauxTest")) Then
585         rstProjet.Fields("TauxTest") = rstSoum.Fields("TauxTest")
590       Else
595         Call rstConfig.Open("SELECT TauxTestMec FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

600         rstProjet.Fields("TauxTest") = rstConfig.Fields("TauxTestMec")

605         Call rstConfig.Close
610       End If

615       If Not IsNull(rstSoum.Fields("TauxInstallation")) Then
620         rstProjet.Fields("TauxInstallation") = rstSoum.Fields("TauxInstallation")
625       Else
630         Call rstConfig.Open("SELECT TauxInstallationMec FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

635         rstProjet.Fields("TauxInstallation") = rstConfig.Fields("TauxInstallationMec")

640         Call rstConfig.Close
645       End If

650       If Not IsNull(rstSoum.Fields("TauxFormation")) Then
655         rstProjet.Fields("TauxFormation") = rstSoum.Fields("TauxFormation")
660       Else
665         Call rstConfig.Open("SELECT TauxFormationMec FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

670         rstProjet.Fields("TauxFormation") = rstConfig.Fields("TauxFormationMec")

675         Call rstConfig.Close
680       End If

685       If Not IsNull(rstSoum.Fields("TauxGestion")) Then
690         rstProjet.Fields("TauxGestion") = rstSoum.Fields("TauxGestion")
695       Else
700         Call rstConfig.Open("SELECT TauxGestionProjetsMec FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

705         rstProjet.Fields("TauxGestion") = rstConfig.Fields("TauxGestionProjetsMec")

710         Call rstConfig.Close
715       End If

720       If Not IsNull(rstSoum.Fields("TauxShipping")) Then
725         rstProjet.Fields("TauxShipping") = rstSoum.Fields("TauxShipping")
730       Else
735         Call rstConfig.Open("SELECT TauxShippingMec FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

740         rstProjet.Fields("TauxShipping") = rstConfig.Fields("TauxShippingMec")

745         Call rstConfig.Close
750       End If

755       Set rstConfig = Nothing

760       rstProjet.Fields("ProchaineCommande") = 1
   
765       Call rstProjet.Update
770     End If

775     Call rstSoum.Close
780     Set rstSoum = Nothing
    
785     Call rstProjet.Close
790     Set rstProjet = Nothing
    
795     Call CalculerTotalRecordset(sNoProjet)

800     Exit Sub

AfficherErreur:

805     woups "frmProjSoumMec", "CreerProjetInstallation", Err, Erl
End Sub

Private Sub cmdDateFacturation_Click()

5       On Error GoTo AfficherErreur

        'Ouverture du calendrier
10      If txtDateFacturation.Text <> vbNullString Then
15        mvwDateFacturation.Value = txtDateFacturation.Text
20      Else
25        mvwDateFacturation.Value = Date
35      End If

40      mvwDateFacturation.Visible = True

45      Call mvwDateFacturation.SetFocus

50      Exit Sub

AfficherErreur:

55      woups "frmProjSoumMec", "cmdDateFacturation_Click", Err, Erl
End Sub

Private Sub cmdDemande_Click()

5       On Error GoTo AfficherErreur

10      Dim rstProjSoum As ADODB.Recordset
15      Dim sUser       As String

20      If Right$(txtNoProjSoum.Text, 2) = "99" Then
25        Call MsgBox("Vous ne pouvez pas commander de pièce à partir de ce projet!", vbOKOnly, "Erreur")

30        Exit Sub
35      End If

40      Set rstProjSoum = New ADODB.Recordset

45      Call rstProjSoum.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

50      If rstProjSoum.Fields("Ouvert") = True And rstProjSoum.Fields("Verrouillé") = False Then
55        If VerifierSiOuvert(sUser) = False Then
60          Call frmChoixDemande.AfficherProjetSoumission(txtNoProjSoum.Text, MECANIQUE, MODE_PIECE, m_eType)
65        Else
70          If m_eType = TYPE_PROJET Then
75            Call MsgBox("Ce projet est en modification par " & sUser & "!", vbOKOnly, "Erreur")
80          Else
85            Call MsgBox("Cette soumission est en modification par " & sUser & "!", vbOKOnly, "Erreur")
90          End If
95        End If
100     Else
105       If rstProjSoum.Fields("Ouvert") = False Then
110         If m_eType = TYPE_PROJET Then
115           Call MsgBox("Ce projet est fermé!", vbOKOnly, "Erreur")
120         Else
125           Call MsgBox("Cette soumission est fermée!", vbOKOnly, "Erreur")
130         End If
135       Else
140         If m_eType = TYPE_PROJET Then
145           Call MsgBox("Ce projet est verrouillé!", vbOKOnly, "Erreur")
150         Else
155           Call MsgBox("Cette soumission est verrouillée!", vbOKOnly, "Erreur")
160         End If
165       End If
170     End If

175     Call rstProjSoum.Close
180     Set rstProjSoum = Nothing

185     Exit Sub

AfficherErreur:

190     woups "frmProjSoumMec", "cmdDemande_Click", Err, Erl
End Sub

Private Sub cmdHistorique_Click()

5       On Error GoTo AfficherErreur

        'Ouverture de l'historique des modifications
10      If cmbProjSoum.ListCount > 0 Then
15        Call RemplirListViewModifications

20        lvwHistorique.Visible = True
  
25        Call lvwHistorique.SetFocus
30      End If

35      Exit Sub

AfficherErreur:

40      woups "frmProjSoumMec", "cmdHistorique_Click", Err, Erl
End Sub

Private Sub cmdLegende_Click()
  
5       On Error GoTo AfficherErreur
  
10      Call OuvrirForm(frmLegende, True)

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMec", "cmdLegende_Click", Err, Erl
End Sub

Private Sub cmdOKFRS_Click()

5       On Error GoTo AfficherErreur

10      If m_bPieceInutile = True Or m_bChangementFRS = True Then
15        Call ChoisirFournisseurMateriel
20      Else
25        Call ChoisirFournisseur
30      End If

35      Exit Sub

AfficherErreur:

40      woups "frmProjSoumMec", "cmdOKFRS_Click", Err, Erl
End Sub

Private Sub cmdOKPrix_Click()
        'Écrit les prix dans le ListView
5       On Error GoTo AfficherErreur

10      Dim rstConfig    As ADODB.Recordset
15      Dim itmSoum      As ListItem
20      Dim itmAvant     As ListItem
25      Dim bPrixSpecial As Boolean
30      Dim iCompteur    As Integer
35      Dim lColor       As Long
40      Dim sPiece       As String
45      Dim sQuantite    As String
50      Dim sTauxUSA     As String
55      Dim sTauxSPA     As String
  
60      Set rstConfig = New ADODB.Recordset

65      Call rstConfig.Open("SELECT TauxAmericain, TauxEspagnol FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)
  
70      sTauxUSA = rstConfig.Fields("TauxAmericain")
75      sTauxSPA = rstConfig.Fields("TauxEspagnol")

80      Call rstConfig.Close
85      Set rstConfig = Nothing
  
90      If m_bMauvaisPrix = False Then
95        If cmbfrs.ListIndex = -1 Then
100         Call MsgBox("Vous devez choisir un fournisseur!", vbOKOnly, "Erreur")
    
105         Exit Sub
110       End If
115     End If
  
120     If Trim$(txtPrixList.Text) = vbNullString Then
125       If Trim$(txtPrixSpecial.Text) = vbNullString Then
130         Call MsgBox("Vous devez remplir le prix listé!", vbOKOnly, "Erreur")

135         Exit Sub
140       End If
145     End If
  
150     If Trim$(txtPrixNet.Text) = vbNullString And Trim$(txtPrixSpecial.Text) = vbNullString Then
155       Call MsgBox("Vous devez choisir un prix!", vbOKOnly, "Erreur")
    
160       Exit Sub
165     Else
170       If Trim$(txtPrixNet.Text) <> "" Then
175         bPrixSpecial = False
180       Else
185         bPrixSpecial = True
190       End If
195     End If

200     If m_bMauvaisPrix = True Then
205       sQuantite = InputBox("Quelle est la quantité!")

210       If sQuantite <> "" Then
215         If Not IsNumeric(sQuantite) Then
220           Exit Sub
225         End If
230       Else
235         Exit Sub
240       End If

245       Set itmAvant = lvwSoumission.SelectedItem
250       Set itmSoum = lvwSoumission.ListItems.Add(CInt(fraPrix.Tag) + 1)
  
255       lColor = itmAvant.ListSubItems(I_COL_SOUM_PIECE).ForeColor
  
260       itmSoum.Checked = itmAvant.Checked
  
          'Quantité
265       itmSoum.Text = "-" & itmAvant.Text

          'On met l'id de la section dans le tag du listItem
270       itmSoum.Tag = itmAvant.Tag
                                                                                                         
          'No d'item
275       itmSoum.SubItems(I_COL_SOUM_PIECE) = itmAvant.SubItems(I_COL_SOUM_PIECE)
   
          'On met le nom de la sous-section dans le tag du no d'item
280       itmSoum.ListSubItems(I_COL_SOUM_PIECE).Tag = itmAvant.ListSubItems(I_COL_SOUM_PIECE).Tag
  
          'On met la description en francais dans la colonne et la description en anglais
          'dans le tag
285       itmSoum.SubItems(I_COL_SOUM_DESCR) = itmAvant.SubItems(I_COL_SOUM_DESCR)
290       itmSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = itmAvant.ListSubItems(I_COL_SOUM_DESCR).Tag

          'On met le nom du fabricant dans la col-nne et l'ordre de la section dans le tag
295       itmSoum.SubItems(I_COL_SOUM_MANUFACT) = itmAvant.SubItems(I_COL_SOUM_MANUFACT)
300       itmSoum.ListSubItems(I_COL_SOUM_MANUFACT).Tag = itmAvant.ListSubItems(I_COL_SOUM_MANUFACT).Tag

          'Prix listé
305       itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = itmAvant.SubItems(I_COL_SOUM_PRIX_LIST)

310       itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = itmAvant.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag
       
315       itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = itmAvant.SubItems(I_COL_SOUM_ESCOMPTE)

320       itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = itmAvant.SubItems(I_COL_SOUM_PRIX_NET)

325       itmSoum.ListSubItems(I_COL_SOUM_PRIX_NET).Tag = itmAvant.ListSubItems(I_COL_SOUM_PRIX_NET).Tag
            
          'On met le fournisseur dans la colonne et l'id dans le tag
330       itmSoum.SubItems(I_COL_SOUM_DISTRIB) = itmAvant.SubItems(I_COL_SOUM_DISTRIB)
335       itmSoum.ListSubItems(I_COL_SOUM_DISTRIB).Tag = itmAvant.ListSubItems(I_COL_SOUM_DISTRIB).Tag
    
          'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
340       itmSoum.SubItems(I_COL_SOUM_TOTAL) = "-" & itmAvant.SubItems(I_COL_SOUM_TOTAL)
      
          'Pour le profit, c'est le prix total - (prix net * quantité)
345       itmSoum.SubItems(I_COL_SOUM_PROFIT) = "-" & itmAvant.SubItems(I_COL_SOUM_PROFIT)

          'Ajout du nouveau Prix
350       Set itmSoum = lvwSoumission.ListItems.Add(CInt(fraPrix.Tag) + 2)

355       itmSoum.Checked = itmAvant.Checked
  
          'Quantité
360       itmSoum.Text = sQuantite

          'On met l'id de la section dans le tag du listItem
365       itmSoum.Tag = itmAvant.Tag
                                                                                                         
          'No d'item
370       itmSoum.SubItems(I_COL_SOUM_PIECE) = itmAvant.SubItems(I_COL_SOUM_PIECE)

375       itmSoum.ListSubItems(I_COL_SOUM_PIECE).ForeColor = lColor
   
          'On met le nom de la sous-section dans le tag du no d'item
380       itmSoum.ListSubItems(I_COL_SOUM_PIECE).Tag = itmAvant.ListSubItems(I_COL_SOUM_PIECE).Tag
  
          'On met la description en francais dans la colonne et la description en anglais
          'dans le tag
385       itmSoum.SubItems(I_COL_SOUM_DESCR) = itmAvant.SubItems(I_COL_SOUM_DESCR)
390       itmSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = itmAvant.ListSubItems(I_COL_SOUM_DESCR).Tag

395       itmSoum.ListSubItems(I_COL_SOUM_DESCR).ForeColor = lColor
          
          'On met le nom du fabricant dans la col-nne et l'ordre de la section dans le tag
400       itmSoum.SubItems(I_COL_SOUM_MANUFACT) = itmAvant.SubItems(I_COL_SOUM_MANUFACT)
405       itmSoum.ListSubItems(I_COL_SOUM_MANUFACT).Tag = itmAvant.ListSubItems(I_COL_SOUM_MANUFACT).Tag

410       itmSoum.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = lColor
          
415       If bPrixSpecial = False Then
420         If optUSA.Value = True Then
425           itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixList.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
430         Else
435           If optSpain.Value = True Then
440             itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixList.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
445           Else
450             itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(txtPrixList.Text, MODE_ARGENT, 4)
455           End If
460         End If

465         itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = Conversion(txtPrixList.Text, MODE_ARGENT, 4)

470         itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = lColor
       
            'Escompte
475         If mskEscompte.Text <> vbNullString Then
480           itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion(mskEscompte.Text, MODE_POURCENT)
485         Else
490           itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion("0", MODE_POURCENT)
495         End If

500         itmSoum.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = lColor
            
            'Prix net
505         If optUSA.Value = True Then
510           itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixNet.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
515         Else
520           If optSpain.Value = True Then
525             itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixNet.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
530           Else
535             itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(txtPrixNet.Text, MODE_ARGENT, 4)
540           End If
545         End If
           
550         itmSoum.ListSubItems(I_COL_SOUM_PRIX_NET).Tag = itmAvant.ListSubItems(I_COL_SOUM_PRIX_NET).Tag

555         itmSoum.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = lColor
560       Else
565         If optUSA.Value = True Then
570           itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
575         Else
580           If optSpain.Value = True Then
585             itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
590           Else
595             itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(txtPrixSpecial.Text, MODE_ARGENT, 4)
600           End If
605         End If

610         itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = Conversion(txtPrixSpecial.Text, MODE_ARGENT, 4)

615         itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = lColor
      
620         itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion("0", MODE_POURCENT)

625         itmSoum.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = lColor

630         If optUSA.Value = True Then
635           itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
640         Else
645           If optSpain.Value = True Then
650             itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
655           Else
660             itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(txtPrixSpecial.Text, MODE_ARGENT, 4)
665           End If
670         End If

675         itmSoum.ListSubItems(I_COL_SOUM_PRIX_NET).Tag = itmAvant.ListSubItems(I_COL_SOUM_PRIX_NET).Tag

680         itmSoum.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = lColor
685       End If

          'On met le fournisseur dans la colonne et l'id dans le tag
690       itmSoum.SubItems(I_COL_SOUM_DISTRIB) = cmbfrs.LIST(cmbfrs.ListIndex)
695       itmSoum.ListSubItems(I_COL_SOUM_DISTRIB).Tag = cmbfrs.ItemData(cmbfrs.ListIndex)

700       itmSoum.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = lColor
          
          'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
705       itmSoum.SubItems(I_COL_SOUM_TOTAL) = Conversion(CStr(Round(Replace(itmSoum.Text, "*", vbNullString) * itmSoum.SubItems(I_COL_SOUM_PRIX_NET) * CSng(m_sProfit), 2)), MODE_ARGENT)

710       itmSoum.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = lColor

715       If optUSA.Value = True Then
720         itmSoum.ListSubItems(I_COL_SOUM_TOTAL).Tag = "USA"
725       Else
730         If optSpain.Value = True Then
735           itmSoum.ListSubItems(I_COL_SOUM_TOTAL).Tag = "SPA"
740         Else
745           itmSoum.ListSubItems(I_COL_SOUM_TOTAL).Tag = "CAN"
750         End If
755       End If
    
          'Pour le profit, c'est le prix total - (prix net * quantité)
760       itmSoum.SubItems(I_COL_SOUM_PROFIT) = Conversion(CStr(Round(itmSoum.SubItems(I_COL_SOUM_TOTAL) - (itmSoum.SubItems(I_COL_SOUM_PRIX_NET) * Replace(itmSoum.Text, "*", "")), 2)), MODE_ARGENT)
          
765       itmSoum.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = lColor

770       itmSoum.SubItems(I_COL_SOUM_COMMENTAIRE) = itmAvant.SubItems(I_COL_SOUM_COMMENTAIRE)
775       itmSoum.ListSubItems(I_COL_SOUM_COMMENTAIRE).ForeColor = lColor

780       If m_eType = TYPE_PROJET Then
785         itmSoum.SubItems(I_COL_SOUM_DATE_COMMANDE) = itmAvant.SubItems(I_COL_SOUM_DATE_COMMANDE)
790         itmSoum.ListSubItems(I_COL_SOUM_DATE_COMMANDE).ForeColor = lColor

795         itmSoum.SubItems(I_COL_SOUM_DATE_REQUISE) = itmAvant.SubItems(I_COL_SOUM_DATE_REQUISE)
800         itmSoum.ListSubItems(I_COL_SOUM_DATE_REQUISE).ForeColor = lColor

805         itmSoum.SubItems(I_COL_SOUM_NOM_COMMANDE) = itmAvant.SubItems(I_COL_SOUM_NOM_COMMANDE)
810         itmSoum.ListSubItems(I_COL_SOUM_NOM_COMMANDE).ForeColor = lColor

815         itmSoum.SubItems(I_COL_SOUM_NO_SEQUENTIEL) = itmAvant.SubItems(I_COL_SOUM_NO_SEQUENTIEL)
820         itmSoum.ListSubItems(I_COL_SOUM_NO_SEQUENTIEL).ForeColor = lColor

825         itmSoum.SubItems(I_COL_SOUM_FACTURATION) = itmAvant.SubItems(I_COL_SOUM_FACTURATION)

830         If itmSoum.SubItems(I_COL_SOUM_FACTURATION) <> "" Then
835           itmSoum.ListSubItems(I_COL_SOUM_FACTURATION).Tag = itmAvant.ListSubItems(I_COL_SOUM_FACTURATION).Tag
840         End If

845         itmSoum.ListSubItems(I_COL_SOUM_FACTURATION).ForeColor = lColor
850       End If

855       If itmAvant.SubItems(I_COL_SOUM_COMMENTAIRE) <> "" Then
860         itmSoum.ListSubItems(I_COL_SOUM_COMMENTAIRE).ForeColor = itmAvant.ListSubItems(I_COL_SOUM_COMMENTAIRE).ForeColor

865         itmAvant.ListSubItems(I_COL_SOUM_COMMENTAIRE).ForeColor = COLOR_NOIR
870       End If

875       If m_eType = TYPE_PROJET Then
880         If itmAvant.SubItems(I_COL_SOUM_DATE_COMMANDE) <> "" Then
885           itmAvant.ListSubItems(I_COL_SOUM_DATE_COMMANDE).ForeColor = COLOR_NOIR
890         End If

895         If itmAvant.SubItems(I_COL_SOUM_DATE_REQUISE) <> "" Then
900           itmAvant.ListSubItems(I_COL_SOUM_DATE_REQUISE).ForeColor = COLOR_NOIR
905         End If

910         If itmAvant.SubItems(I_COL_SOUM_NOM_COMMANDE) <> "" Then
915           itmAvant.ListSubItems(I_COL_SOUM_NOM_COMMANDE).ForeColor = COLOR_NOIR
920         End If

925         If itmAvant.SubItems(I_COL_SOUM_NO_SEQUENTIEL) <> "" Then
930           itmAvant.ListSubItems(I_COL_SOUM_NO_SEQUENTIEL).ForeColor = COLOR_NOIR
935         End If

940         itmAvant.ListSubItems(I_COL_SOUM_FACTURATION).ForeColor = COLOR_NOIR

945         If itmAvant.SubItems(I_COL_SOUM_PROVENANCE) <> "" Then
950           itmAvant.ListSubItems(I_COL_SOUM_PROVENANCE).ForeColor = COLOR_NOIR
955         End If
960       End If

965       itmSoum.ListSubItems(I_COL_SOUM_DESCR).ForeColor = itmAvant.ListSubItems(I_COL_SOUM_DESCR).ForeColor
970       itmSoum.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = itmAvant.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor
975       itmSoum.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = itmAvant.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor
980       itmSoum.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = itmAvant.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor
985       itmSoum.ListSubItems(I_COL_SOUM_PIECE).ForeColor = itmAvant.ListSubItems(I_COL_SOUM_PIECE).ForeColor
990       itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = itmAvant.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor
995       itmSoum.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = itmAvant.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor
1000      itmSoum.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = itmAvant.ListSubItems(I_COL_SOUM_PROFIT).ForeColor
1005      itmSoum.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = itmAvant.ListSubItems(I_COL_SOUM_TOTAL).ForeColor

1010      itmAvant.ListSubItems(I_COL_SOUM_DESCR).ForeColor = COLOR_NOIR
1015      itmAvant.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = COLOR_NOIR
1020      itmAvant.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = COLOR_NOIR
1025      itmAvant.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = COLOR_NOIR
1030      itmAvant.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_NOIR
1035      itmAvant.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = COLOR_NOIR
1040      itmAvant.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = COLOR_NOIR
1045      itmAvant.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = COLOR_NOIR
1050      itmAvant.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = COLOR_NOIR

1055      If lvwSoumission.ListItems.count > 0 Then
1060        Call Deselect

1065        lvwSoumission.ListItems(1).Selected = True
1070      End If

1075      m_bMauvaisPrix = False

1080      cmbfrs.Locked = False

1085      Call lvwSoumission.Refresh
1090    Else
1095      sPiece = lvwSoumission.ListItems(CInt(fraPrix.Tag)).SubItems(I_COL_SOUM_PIECE)

1100      For iCompteur = 1 To lvwSoumission.ListItems.count
1105        If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) = sPiece And lvwSoumission.ListItems(iCompteur).ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_MAGENTA Then
1110          Set itmSoum = lvwSoumission.ListItems(iCompteur)

1115          itmSoum.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_NOIR
1120          itmSoum.ListSubItems(I_COL_SOUM_DESCR).ForeColor = COLOR_NOIR
1125          itmSoum.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = COLOR_NOIR
1130          itmSoum.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = COLOR_NOIR
1135          itmSoum.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = COLOR_NOIR
1140          itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = COLOR_NOIR
1145          itmSoum.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = COLOR_NOIR
1150          itmSoum.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = COLOR_NOIR
1155          itmSoum.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = COLOR_NOIR

1160          If itmSoum.SubItems(I_COL_SOUM_COMMENTAIRE) <> "" Then
1165            itmSoum.ListSubItems(I_COL_SOUM_COMMENTAIRE).ForeColor = COLOR_NOIR
1170          End If

1175          Call lvwSoumission.Refresh
           
1180          If bPrixSpecial = False Then
                'Prix listé
1185            If optUSA.Value = True Then
1190              itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixList.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
1195            Else
1200              If optSpain.Value = True Then
1205                itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixList.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
1210              Else
1215                itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(txtPrixList.Text, MODE_ARGENT, 4)
1220              End If
1225            End If
               
1230            itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = Conversion(txtPrixList.Text, MODE_DECIMAL, 4)

                'Escompte
1235            If mskEscompte.Text <> vbNullString Then
1240              itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion(mskEscompte.Text, MODE_POURCENT)
1245            Else
1250              itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion("0", MODE_POURCENT)
1255            End If
  
                'Prix net
1260            If optUSA.Value = True Then
1265              itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixNet.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
1270            Else
1275              If optSpain.Value = True Then
1280                itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixNet.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
1285              Else
1290                itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(txtPrixNet.Text, MODE_ARGENT, 4)
1295              End If
1300            End If
1305          Else
                'Prix listé
1310            If optUSA.Value = True Then
1315              itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
1320            Else
1325              If optSpain.Value = True Then
1330                itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
1335              Else
1340                itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(txtPrixSpecial.Text, MODE_ARGENT, 4)
1345              End If
1350            End If
                
1355            itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = Conversion(txtPrixSpecial.Text, MODE_DECIMAL, 4)
      
1360            itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion("0", MODE_POURCENT)
  
                'Prix net
1365            If optUSA.Value = True Then
1370              itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text), 4)), MODE_ARGENT, 4)
1375            Else
1380              If optSpain.Value = True Then
1385                itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text), 4)), MODE_ARGENT, 4)
1390              Else
1395                itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(txtPrixSpecial.Text, MODE_ARGENT, 4)
1400              End If
1405            End If
1410          End If
         
              'On met le fournisseur dans la colonne et l'id dans le tag
1415          itmSoum.SubItems(I_COL_SOUM_DISTRIB) = cmbfrs.LIST(cmbfrs.ListIndex)
  
1420          itmSoum.ListSubItems(I_COL_SOUM_DISTRIB).Tag = cmbfrs.ItemData(cmbfrs.ListIndex)
          
              'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
1425          itmSoum.SubItems(I_COL_SOUM_TOTAL) = Conversion(CStr(Round(itmSoum.Text * Conversion(itmSoum.SubItems(I_COL_SOUM_PRIX_NET), MODE_PAS_FORMAT) * CSng(m_sProfit), 2)), MODE_ARGENT)

1430          If optUSA.Value = True Then
1435            itmSoum.ListSubItems(I_COL_SOUM_TOTAL).Tag = "USA"
1440          Else
1445            If optSpain.Value = True Then
1450              itmSoum.ListSubItems(I_COL_SOUM_TOTAL).Tag = "SPA"
1455            Else
1460              itmSoum.ListSubItems(I_COL_SOUM_TOTAL).Tag = "CAN"
1465            End If
1470          End If
      
              'Pour le profit, c'est le prix total - (prix net * quantité)
1475          itmSoum.SubItems(I_COL_SOUM_PROFIT) = Conversion(CStr(Round(itmSoum.SubItems(I_COL_SOUM_TOTAL) - (itmSoum.SubItems(I_COL_SOUM_PRIX_NET) * itmSoum.Text), 2)), MODE_ARGENT)
1480        End If
1485      Next
1490    End If

1495    Call ModifierPrixCatalogue

1500    fraPrix.Visible = False

1505    Call CalculerPrix

1510    Exit Sub

AfficherErreur:

1515    woups "frmProjSoumMec", "cmdOKPrix_Click", Err, Erl
End Sub

Private Sub cmdPhotos_Click()

5       On Error GoTo AfficherErreur

10      If txtCheminPhotos.Text <> vbNullString Then
15        Call frmPhotoProjSoum.Afficher(txtCheminPhotos.Text)
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmProjSoumMec", "cmdPhotos_Click", Err, Erl
End Sub

Private Sub cmdReception_Click()

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer
15      Dim bOuvert   As Boolean

20      If Right$(txtNoProjSoum.Text, 2) = "99" Then
25        Call MsgBox("Vous ne pouvez pas faire de réception pour ce projet!", vbOKOnly, "Erreur")

30        Exit Sub
35      End If

40      For iCompteur = 0 To Forms.count - 1
45        If Forms(iCompteur).Name = "FrmReceptionMec" Then
50          bOuvert = True

55          Exit For
60        End If
65      Next

70      If bOuvert = True Then
75        Call Unload(FrmReceptionMec)
80      End If

85      Call FrmReceptionMec.AfficherProjet(g_sUserID, txtNoProjSoum.Text)

90      Call RemplirListViewProjSoum(txtNoProjSoum.Text)

95      Exit Sub

AfficherErreur:

100     woups "frmProjSoumMec", "cmdReception_Click", Err, Erl
End Sub

Private Sub cmdRechercherClient_Click()

5       On Error GoTo AfficherErreur

10      Dim sRecherche As String
    
15      sRecherche = InputBox("Entrez le texte à rechercher.")
     
20      If StrPtr(sRecherche) <> 0 Then
25        Call RemplirComboClients(sRecherche)
30      End If
  
35      Exit Sub

AfficherErreur:

40      woups "frmProjSoumMec", "cmdRechercherClient_Click", Err, Erl
End Sub

Private Sub cmdRetour_Click()

5       On Error GoTo AfficherErreur

10      If Right$(txtNoProjSoum.Text, 2) = "99" Then
15        Call MsgBox("Vous ne pouvez pas faire de retour dans ce projet!", vbOKOnly, "Erreur")

20        Exit Sub
25      End If

30      Screen.MousePointer = vbHourglass

35      Call frmRetourMarchandise.Afficher(txtNoProjSoum.Text, g_sUserID)

40      Call cmbProjSoum_Click

45      Screen.MousePointer = vbDefault

50      Exit Sub

AfficherErreur:

55      woups "frmProjSoumMec", "cmdRetour_Click", Err, Erl
End Sub

Private Sub cmdBrowse_Click()

5       On Error GoTo AfficherErreur

10      Call frmChoixDossier.Afficher(Me)

15      If m_bAnnulerChemin = False Then
20        txtCheminPhotos.Text = m_sChemin
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmProjSoumMec", "cmdBrowse_Click", Err, Erl
End Sub

Private Sub cmdRafraichir_Click()

5       On Error GoTo AfficherErreur

10      If m_sTri <> vbNullString Then
15        m_sTri = vbNullString
  
20        Call RemplirListViewPieces
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmProjSoumMec", "cmdRafraichir_Click", Err, Erl
End Sub

Private Sub cmdRapportFACT_Click()

5       On Error GoTo AfficherErreur

10      Dim rstProjSoum As ADODB.Recordset
15      Dim sUser       As String
20      Dim sNoFacture  As String

25      If lvwSoumission.ListItems.count > 0 Then
30        If txtNoProjSoum.Text <> vbNullString Then
35          If VerifierSiOuvert(sUser) = False Then
40            sNoFacture = lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_FACTURATION)

45            If Left$(sNoFacture, 2) = "F-" Or sNoFacture = "NC" Then
                'Ouvre les tables
50              Set rstProjSoum = New ADODB.Recordset
                
55              Call rstProjSoum.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = '" & txtNoProjSoum.Text & "' ORDER BY IDProjet", g_connData, adOpenDynamic, adLockOptimistic)

                '***********************************************************************************
                'AJOUT PAR GAÉTAN GINGRAS 07 FÉVRIER 2010
                '***********************************************************************************
                If MsgBox("Désirez-vous afficher les dates de réception et de commande?", vbYesNo, "Date de réception et de commande") = vbYes Then
                    bFlag = True
                Else
                    bFlag = False
                End If
                '***********************************************************************************

60              Call ImprimerProjSoumFacturation(rstProjSoum, sNoFacture)
65              Call ImprimerListePiecesFacturation(rstProjSoum, sNoFacture)

70              'Call rstProjSoum.MoveNext

75              Call rstProjSoum.Close
80              Set rstProjSoum = Nothing
85            Else
90              Call MsgBox("La ligne sélectionnée ne contient aucune facture!", vbOKOnly, "Erreur")
95            End If
100         Else
105           Call MsgBox("Ce projet est en modification par " & sUser & "!", vbOKOnly, "Erreur")
110         End If
115       End If
120     Else
125       Call MsgBox("Ce projet ne contient aucune pièce à imprimer!", vbOKOnly, "Erreur")
130     End If

135     Exit Sub

AfficherErreur:

140     woups "frmProjSoumMec", "cmdRapportFACT_Click", Err, Erl
End Sub

Private Sub cmdSortieMagasin_Click()

5       On Error GoTo AfficherErreur

10      Call SortieMagasin

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMec", "cmdSortieMagasin_Click", Err, Erl
End Sub

Private Sub ChangerQuantite()

5       On Error GoTo AfficherErreur

10      Dim sQuantite As String
15      Dim itmSoum   As ListItem

20      sQuantite = Replace(InputBox("Quelle est la nouvelle quantité?"), ".", ",")

25      If IsNumeric(sQuantite) Then
30        Set itmSoum = lvwSoumission.SelectedItem

35        itmSoum.Text = sQuantite

          'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
40        itmSoum.SubItems(I_COL_SOUM_TOTAL) = Conversion(Round(Replace(itmSoum.Text, "*", vbNullString) * itmSoum.SubItems(I_COL_SOUM_PRIX_NET) * CSng(m_sProfit), 2), MODE_ARGENT, 2)
      
          'Pour le profit, c'est le prix total - (prix net * quantité)
45        itmSoum.SubItems(I_COL_SOUM_PROFIT) = Conversion(CStr(Round(itmSoum.SubItems(I_COL_SOUM_TOTAL) - (itmSoum.SubItems(I_COL_SOUM_PRIX_NET) * Replace(itmSoum.Text, "*", vbNullString)), 2)), MODE_ARGENT)

50        Call CalculerPrix
55      End If

60      Exit Sub

AfficherErreur:

65      woups "frmProjSoumMec", "ChangerQuantite", Err, Erl
End Sub

Private Sub SortieMagasin()

5       On Error GoTo AfficherErreur

10      Dim lColor As Long
15      Dim sTag   As String

20      If lvwSoumission.ListItems.count > 0 Then
25        If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PROFIT).Tag <> "EXTRA" Then
            'Si pas une section
30          If lvwSoumission.SelectedItem.Tag <> "" Then
              'Si pas une sous-section
35            If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) <> "" Then
                'Si ce n'est pas du Texte
40              If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) <> "Texte" And lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) <> "Text" Then
                  'Si la pièce est noire ou gris
45                If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_NOIR Or lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_GRIS Then
                    'Si la pièce est noire
50                  If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_NOIR Then
                      'On la met grise
55                    lColor = COLOR_ORANGE
  
60                    sTag = Replace(lvwSoumission.SelectedItem.Text, "*", "")
65                  Else
70                    lColor = COLOR_NOIR

75                    sTag = "0"
80                  End If
                
85                  lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_DESCR).ForeColor = lColor
90                  lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = lColor
95                  lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = lColor
100                 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = lColor
105                 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor = lColor
110                 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = lColor
115                 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = lColor
120                 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = lColor
125                 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = lColor

130                 If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_FACTURATION) = "" Then
135                   lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_FACTURATION) = " "
140                 End If

145                 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_FACTURATION).Tag = sTag

150                 Call lvwSoumission.Refresh

155                 Call CalculerPrixReception
160               End If
165             End If
170           End If
175         End If
180       Else
185         Call MsgBox("Cette commande doit être faite dans le projet " & lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PROVENANCE), vbOKOnly, "Erreur")
190       End If
195     End If

200     Exit Sub

AfficherErreur:

205     woups "frmProjSoumMec", "SortieMagasin", Err, Erl
End Sub

Private Sub cmdSupprimerFRS_Click()
        'Permet d'effacer un Fournisseur
5       On Error GoTo AfficherErreur

10      Dim sPiece As String

        'Si c'est pas Choisir Ultérieurement
15      If lvwfournisseur.SelectedItem.Text <> "CHOISIR ULTÉRIEUREMENT" Then
20        If m_bPieceInutile = True Then
25          sPiece = lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE)
30        Else
35          If m_bRecherchePiece = True Then
40            sPiece = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM)
45          Else
50            sPiece = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_NO_ITEM)
55          End If
60        End If

65        If MsgBox("Voulez-vous vraiment supprimer le fournisseur " & lvwfournisseur.SelectedItem.Text & " pour la pièce " & lvwPieces.SelectedItem.SubItems(I_COL_PIECES_NO_ITEM) & "?", vbYesNo, "Suppression") = vbYes Then
70          Call g_connData.Execute("DELETE * FROM GRB_PiecesFRS WHERE NoEnreg = " & lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_ENTRER_PAR).Tag)

75          Call RemplirListViewFournisseur

80          frafournisseur.Visible = True
85          Call lvwfournisseur.SetFocus
90        End If
95      End If

100     Exit Sub

AfficherErreur:

105     woups "frmProjSoumMec", "cmdSupprimerFRS_Click", Err, Erl
End Sub

Private Sub cmdTemps_Click()

5       On Error GoTo AfficherErreur

10      If cmbProjSoum.ListCount > 0 Then
15        If m_eMode = MODE_AJOUT_MODIF Then
20          If m_bModeAjout = True Then
25            If m_bExtra = True Then
30              If m_eType = TYPE_PROJET Then
35                Call frmProjSoumMecTemps.Afficher(txtNoProjSoum.Text, txtNoSoumission.Text, m_eType, m_eMode, False)
40              Else
45                Call frmProjSoumMecTemps.Afficher("", txtNoProjSoum.Text, m_eType, m_eMode, False)
50              End If
55            Else
60              If m_eType = TYPE_PROJET Then
65                Call frmProjSoumMecTemps.Afficher(txtNoProjSoum.Text, txtNoSoumission.Text, m_eType, m_eMode, True)
70              Else
75                Call frmProjSoumMecTemps.Afficher("", txtNoProjSoum.Text, m_eType, m_eMode, True)
80              End If
85            End If
90          Else
95            If m_eType = TYPE_PROJET Then
100             Call frmProjSoumMecTemps.Afficher(txtNoProjSoum.Text, txtNoSoumission.Text, m_eType, m_eMode, False)
105           Else
110             Call frmProjSoumMecTemps.Afficher("", txtNoProjSoum.Text, m_eType, m_eMode, False)
115           End If
120         End If
125       Else
130         If m_eType = TYPE_PROJET Then
135           Call frmProjSoumMecTemps.Afficher(txtNoProjSoum.Text, txtNoSoumission.Text, m_eType, m_eMode, False)
140         Else
145           Call frmProjSoumMecTemps.Afficher("", txtNoProjSoum.Text, m_eType, m_eMode, False)
150         End If
155       End If
160     End If

165     If m_eMode = MODE_AJOUT_MODIF Then
170       Call CalculerPrix
175     End If
  
180     m_bTempsDejaOuvert = True

185     Exit Sub

AfficherErreur:

190     woups "frmProjSoumMec", "cmdTemps_Click", Err, Erl
End Sub

Private Sub cmdTexte_Click()

5       On Error GoTo AfficherErreur

10      Dim iIndex       As Integer
15      Dim sSousSection As String
20      Dim sTexte       As String

        'Ajout de texte dans la soumission
25      If lvwSoumission.ListItems.count > 0 Then
30        If lvwSoumission.SelectedItem.Index = 1 Then
35          sSousSection = InputBox("Quelle est la sous-section?")

40          If Trim$(sSousSection) = "" Then
45            sSousSection = S_PAS_SOUS_SECTION
50          End If

55          sTexte = InputBox("Quel est le texte?")

60          If Trim$(sTexte) <> "" Then
65            If Len(sTexte) > 255 Then
70              Call MsgBox("Le texte ne doit pas dépasser 255 caractères!", vbOKOnly, "Erreur")
75            Else
80              iIndex = TrouverIndexSection(sSousSection)

85              Call AjouterTexte(iIndex, sTexte, sSousSection)
90            End If
95          End If
100       Else
105         sTexte = InputBox("Quel est le texte?")

110         If Trim$(sTexte) <> "" Then
115           If Len(sTexte) > 255 Then
120             Call MsgBox("Le texte ne doit pas dépasser 255 caractères!", vbOKOnly, "Erreur")
125           Else
130             iIndex = lvwSoumission.SelectedItem.Index

135             Call AjouterTexte(iIndex, sTexte, "")
140           End If
145         End If
150       End If
155     End If

160     Exit Sub

AfficherErreur:

165     woups "frmProjSoumMec", "cmdTexte_Click", Err, Erl
End Sub

Private Sub AjouterTexte(ByVal iIndex As Integer, ByVal sTexte As String, ByVal sNomSousSection As String)

5       On Error GoTo AfficherErreur

        'Méthode pour ajouter le texte
10      Dim sSousSection As String
15      Dim sOrdre       As String
20      Dim sIDSection   As String
  
        'Si il faut l'ajouter à la fin, on prend les infos du dernier enregistrement
25      If iIndex > lvwSoumission.ListItems.count Then
30        If sNomSousSection = "" Then
35          sSousSection = lvwSoumission.ListItems(iIndex - 1).ListSubItems(I_COL_SOUM_PIECE).Tag
40        Else
45          sSousSection = sNomSousSection
50        End If

55        sOrdre = lvwSoumission.ListItems(iIndex - 1).ListSubItems(I_COL_SOUM_MANUFACT).Tag
60        sIDSection = lvwSoumission.ListItems(iIndex - 1).Tag
65      Else
          'Si c'est une section
70        If lvwSoumission.ListItems(iIndex).Tag = vbNullString Then
75          If sNomSousSection = "" Then
80            sSousSection = lvwSoumission.ListItems(iIndex - 1).ListSubItems(I_COL_SOUM_PIECE).Tag
85          Else
90            sSousSection = sNomSousSection
95          End If

100         sOrdre = lvwSoumission.ListItems(iIndex - 1).ListSubItems(I_COL_SOUM_MANUFACT).Tag
105         sIDSection = lvwSoumission.ListItems(iIndex - 1).Tag
110       Else
            'Si c'est une sous-section
115         If lvwSoumission.ListItems(iIndex).SubItems(I_COL_SOUM_PIECE) = vbNullString Then
              'Si c'est pas la première sous-section
120           If lvwSoumission.ListItems(iIndex - 1).Tag <> vbNullString Then
125             If sNomSousSection = "" Then
130               sSousSection = lvwSoumission.ListItems(iIndex - 1).ListSubItems(I_COL_SOUM_PIECE).Tag
135             Else
140               sSousSection = sNomSousSection
145             End If

150             sOrdre = lvwSoumission.ListItems(iIndex - 1).ListSubItems(I_COL_SOUM_MANUFACT).Tag
155             sIDSection = lvwSoumission.ListItems(iIndex).Tag
160           Else
165             Call MsgBox("Vous ne pouvez pas mettre du texte entre une section et une sous-section!", vbOKOnly, "Erreur")
        
170             Exit Sub
175           End If
180         Else
185           If sNomSousSection = "" Then
190             sSousSection = lvwSoumission.ListItems(iIndex).ListSubItems(I_COL_SOUM_PIECE).Tag
195           Else
200             sSousSection = sNomSousSection
205           End If

210           sOrdre = lvwSoumission.ListItems(iIndex).ListSubItems(I_COL_SOUM_MANUFACT).Tag
215           sIDSection = lvwSoumission.ListItems(iIndex).Tag
220         End If
225       End If
230     End If
    
235     Call lvwSoumission.ListItems.Add(iIndex)
  
240     Call ValeurParDefaut(lvwSoumission.ListItems(iIndex))
  
245     If m_eLangage = ANGLAIS Then
250       lvwSoumission.ListItems(iIndex).SubItems(I_COL_SOUM_PIECE) = "Text"
255     Else
260       lvwSoumission.ListItems(iIndex).SubItems(I_COL_SOUM_PIECE) = "Texte"
265     End If

270     lvwSoumission.ListItems(iIndex).SubItems(I_COL_SOUM_DESCR) = sTexte
  
        'ID de la section
275     lvwSoumission.ListItems(iIndex).Tag = sIDSection
  
        'On ne peut pas écrire dans le tag si c'est vide
280     lvwSoumission.ListItems(iIndex).SubItems(I_COL_SOUM_MANUFACT) = " "
  
        'Ordre de la section
285     lvwSoumission.ListItems(iIndex).ListSubItems(I_COL_SOUM_MANUFACT).Tag = sOrdre
  
        'Nom de la sous-section
290     lvwSoumission.ListItems(iIndex).ListSubItems(I_COL_SOUM_PIECE).Tag = sSousSection

295     Exit Sub

AfficherErreur:

300     woups "frmProjSoumMec", "AjouterTexte", Err, Erl
End Sub

Private Sub cmdTri_Click()

5       On Error GoTo AfficherErreur

10      m_sTri = InputBox("Quel est le texte à trier?")
  
15      m_iCol = cmbTri.ListIndex
  
20      If m_sTri <> vbNullString Then
25        Call RemplirListViewPieces
30      End If

35      Exit Sub

AfficherErreur:

40      woups "frmProjSoumMec", "cmdTri_Click", Err, Erl
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

5       On Error GoTo AfficherErreur

10      If m_eMode = MODE_AJOUT_MODIF Then
15        Call MsgBox("Veuillez enregistrer ou annuler avant de fermer!", vbOKOnly, "Erreur")

20        Cancel = 1
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmProjSoumMec", "Form_QueryUnload", Err, Erl
End Sub

Private Sub Form_Resize()

5       On Error GoTo AfficherErreur
    
10      If Me.Height > I_HEIGHT_AFFICHAGE Then
15        If m_eMode = MODE_INACTIF Then
20          lvwPieces.width = Me.width - 375
25          lvwSoumission.width = Me.width - 375
  
30          lvwSoumission.Height = Me.Height - I_HEIGHT_AFFICHAGE
35          lvwPieces.Height = lvwSoumission.Height * 0.49
40        End If
  
45        cmdImprimer.Top = Me.Height - 825
50        Cmdajouter.Top = Me.Height - 825
55        cmdModifier.Top = Me.Height - 825
60        cmdsupprimer.Top = Me.Height - 825
65        Cmdfermer.Top = Me.Height - 825
70        cmdEnregistrer.Top = Me.Height - 825
75        cmdAnnuler.Top = Me.Height - 825
80        cmdTexte.Top = Me.Height - 825
85        cmdCreerProjet.Top = Me.Height - 825
90        cmdCopier.Top = Me.Height - 825
95        cmdRetour.Top = Me.Height - 825
100       cmdBonCommande.Top = Me.Height - 825
105       cmdDemande.Top = Me.Height - 825
110       cmdAnglaisFrancais.Top = Me.Height - 825
115       cmdExtra.Top = Me.Height - 825
120       cmdCatalogue.Top = Me.Height - 825
125       cmdMaterielInutile.Top = Me.Height - 825
130       cmdReset.Top = Me.Height - 825
135       cmdMauvaisPrix.Top = Me.Height - 825
140       cmdRapportFACT.Top = Me.Height - 825
145       cmdSortieMagasin.Top = Me.Height - 825
150       cmdReception.Top = Me.Height - 825
155     End If

160     Call PositionnerBoutons

165     Exit Sub

AfficherErreur:

170     woups "frmProjSoumMec", "Form_Resize", Err, Erl
End Sub

Private Sub PositionnerBoutons()

5       On Error GoTo AfficherErreur

10      Cmdfermer.Left = Me.width - 1230
15      cmdModifier.Left = Me.width - 2310
20      cmdAnnuler.Left = Me.width - 2310
25      cmdEnregistrer.Left = Me.width - 3390
30      cmdCatalogue.Left = Me.width - 6630
 
35      If m_eType = TYPE_PROJET Then
40        cmdMaterielInutile.Left = Me.width - 7710
45        cmdMauvaisPrix.Left = Me.width - 8790
50        cmdSortieMagasin.Left = Me.width - 9870

55        If m_bSupprimer = True Then
60          cmdsupprimer.Left = Me.width - 3390
65          Cmdajouter.Left = Me.width - 4470
70          cmdBonCommande.Left = Me.width - 5550
75          cmdDemande.Left = Me.width - 6630
80          cmdExtra.Left = Me.width - 7710
85          cmdRetour.Left = Me.width - 8790
90          cmdReception.Left = Me.width - 9870
95        Else
100         Cmdajouter.Left = Me.width - 3390
105         cmdBonCommande.Left = Me.width - 4470
110         cmdDemande.Left = Me.width - 5550
115         cmdExtra.Left = Me.width - 6630
120         cmdRetour.Left = Me.width - 7710
125         cmdReception.Left = Me.width - 8790
130       End If
135     Else
140       cmdsupprimer.Left = Me.width - 3390
145       Cmdajouter.Left = Me.width - 4470
150       cmdCopier.Left = Me.width - 5550
155       cmdMauvaisPrix.Left = Me.width - 7710
160       cmdDemande.Left = Me.width - 6630
165       cmdCreerProjet.Left = Me.width - 7710
170     End If

175     Exit Sub

AfficherErreur:

180     woups "frmProjSoumMec", "PositionnerBoutons", Err, Erl
End Sub

Private Sub Form_Unload(Cancel As Integer)

5       On Error GoTo AfficherErreur

10      Set FrmProjSoumMec = Nothing

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMec", "Form_Unload", Err, Erl
End Sub

Private Sub lvwfournisseur_KeyDown(KeyCode As Integer, Shift As Integer)
        'Permet d'effacer un Fournisseur
5       On Error GoTo AfficherErreur

10      Dim sPiece As String

15      If KeyCode = vbKeyDelete Then
20        If g_bModificationCatalogueMec = True Then
            'Si c'est pas Choisir ultérieurement
25          If lvwfournisseur.SelectedItem.Text <> "CHOISIR ULTÉRIEUREMENT" Then
30            If m_bPieceInutile = True Then
35              sPiece = lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE)
40            Else
45              If m_bRecherchePiece = True Then
50                sPiece = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM)
55              Else
60                sPiece = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_NO_ITEM)
65              End If
70            End If

75            If MsgBox("Voulez-vous vraiment supprimer le fournisseur " & lvwfournisseur.SelectedItem.Text & " pour la pièce " & sPiece & "?", vbYesNo, "Suppression") = vbYes Then
80              Call g_connData.Execute("DELETE * FROM GRB_PiecesFRS WHERE NoEnreg = " & lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_ENTRER_PAR).Tag)

85              Call RemplirListViewFournisseur

90              frafournisseur.Visible = True
95              Call lvwfournisseur.SetFocus
100           End If
105         End If
110       End If
115     End If

120     Exit Sub

AfficherErreur:

125     woups "frmProjSoumMec", "lvwFournisseur_KeyDown", Err, Erl
End Sub

Private Sub lvwHistorique_LostFocus()

5       On Error GoTo AfficherErreur

        'Lorsque l'historique perd le focus, on l'enlève
10      lvwHistorique.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMec", "lvwHistorique_LostFocus", Err, Erl
End Sub

Private Sub lvwPieces_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

5       On Error GoTo AfficherErreur

10      Dim sTexte As String
  
15      sTexte = InputBox("Quel est le texte à rechercher?")
  
20      If Trim$(sTexte) <> vbNullString Then
25        If Len(Trim$(sTexte)) >= 2 Then
30          Call RemplirListViewRecherche(ColumnHeader.Index - 1, sTexte)

35          If lvwPieceTrouve.ListItems.count > 0 Then
40            fraPieceTrouve.Visible = True
45          Else
50            Call MsgBox("Aucun enregistrements trouvés!", vbOKOnly, "Erreur")
55          End If
60        Else
65          Call MsgBox("Il faut un minimum de 2 caractères pour rechercher!", vbOKOnly, "Erreur")
70        End If
75      End If

80      Exit Sub

AfficherErreur:

85      woups "frmProjSoumMec", "lvwPieces_ColumnClick", Err, Erl
End Sub

Private Sub cmdEnregistrer_Click()

5       On Error GoTo AfficherErreur

10      Dim objControl As Object
15      Dim sMessage   As String

20      frafournisseur.Visible = False
25      fraPieceTrouve.Visible = False
30      fraCommentaire.Visible = False
35      fraDateRequise.Visible = False
  
        'Vérification des textbox
40      For Each objControl In Me
45        If TypeOf objControl Is TextBox Then
50          If objControl.Visible = True Then
55            If objControl.Name <> "txtNoSoumission" And _
                 objControl.Name <> "txtCheminPhotos" And _
                 objControl.Name <> "txtPrixReception" And _
                 objControl.Name <> "txtDateFacturation" And _
                 objControl.Name <> "txtPrixSoumission" And _
                 objControl.Name <> "txtForfait" Then
60              If Trim$(objControl.Text) = vbNullString Then
65                Call MsgBox("Vous devez remplir tous les champs!", vbOKOnly, "Erreur")
        
70                Exit Sub
75              End If
80            End If
85          End If
90        Else
95          If TypeOf objControl Is ComboBox Then
100           If objControl.Visible = True Then
105             If objControl.ListIndex = -1 Then
110               If objControl.Name <> "cmbTri" And _
                     objControl.Name <> "cmbSections" And _
                     objControl.Name <> "cmbPieces" Then
115                 Call MsgBox("Vous devez remplir tous les champs!", vbOKOnly, "Erreur")

120                 Exit Sub
125               End If
130             End If
135           End If
140         End If
145       End If
150     Next

155     Screen.MousePointer = vbHourglass
                
160     If BackupPieces(txtNoProjSoum.Text) = False Then
165       If m_eType = TYPE_PROJET Then
170         sMessage = "Une erreur est survenue lors de la copie de sauvegarde du projet en cours!"
175       Else
180         sMessage = "Une erreur est survenue lors de la copie de sauvegarde de la soumission en cours!"
185       End If

190       sMessage = sMessage & vbNewLine & vbNewLine & "Voulez-vous continuer ?"

195       Screen.MousePointer = vbDefault

200       If MsgBox(sMessage, vbYesNo) = vbNo Then
205         Exit Sub
210       Else
215         Screen.MousePointer = vbHourglass
220       End If
225     End If

        'Enregistre la soumission
230     Call EnregistrerProjSoum(txtNoProjSoum.Text)
  
235     Call OuvrirProjSoum(False)
  
        'Remet en mode inactif
240     Call AfficherControles(MODE_INACTIF)
  
245     m_bEnregistrement = True
  
        'Affiche la soumission actuel
250     Call AfficherProjSoum(txtNoProjSoum.Text)

255     m_bEnregistrement = False
  
260     Screen.MousePointer = vbDefault

265     Exit Sub

AfficherErreur:

270     woups "frmProjSoumMec", "cmdEnregistrer_Click", Err, Erl
End Sub

Private Sub EnregistrerFACT(ByVal sNoProjet As String)
        'Calcul le total de chaque facture dans le projet
5       On Error GoTo AfficherErreur

10      Dim rstModif      As ADODB.Recordset
15      Dim rstEmploye    As ADODB.Recordset
20      Dim sPrixTotal    As String
25      Dim sProfit       As String
30      Dim sCommission   As String
35      Dim sNoFacture    As String
40      Dim sTotalPiece   As String
45      Dim sImprevue     As String
50      Dim sTotalTemps   As String
55      Dim sAutre        As String
60      Dim iCompteur     As Integer
65      Dim iIndexFacture As Integer
70      Dim collFacture   As Collection
75      Dim bExiste       As Boolean

80      Set collFacture = New Collection

85      Call g_connData.Execute("DELETE * FROM GRB_Projet_Modif WHERE IDProjet = '" & sNoProjet & "' AND Type = 'M' AND TypeModif = 'FACTURATION'")

90      If lvwSoumission.ListItems.count > 0 Then
95        For iCompteur = 1 To lvwSoumission.ListItems.count
100         bExiste = False

105         sNoFacture = lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_FACTURATION)

110         If sNoFacture <> "" Then
115           For iIndexFacture = 1 To collFacture.count
120             If collFacture(iIndexFacture) = sNoFacture Then
125               bExiste = True

130               Exit For
135             End If
140           Next

145           If bExiste = False Then
150             Call collFacture.Add(sNoFacture)

155             Call CalculerPrixFacturation(sNoFacture, sCommission, sPrixTotal, sProfit, sTotalPiece, sImprevue, sTotalTemps, sAutre)

160             Set rstModif = New ADODB.Recordset

165             Call rstModif.Open("SELECT * FROM GRB_Projet_Modif WHERE Date = '" & Replace(sNoFacture, "F-", "") & "' AND TypeModif = 'FACTURATION'", g_connData, adOpenDynamic, adLockOptimistic)

170             If rstModif.EOF Then
175               Call rstModif.AddNew
180             End If

185             rstModif.Fields("IDProjet") = txtNoProjSoum.Text

190             Set rstEmploye = New ADODB.Recordset

195             Call rstEmploye.Open("SELECT * FROM GRB_Employés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)

200             rstModif.Fields("NoEmployé") = rstEmploye.Fields("NoEmploye")

205             Call rstEmploye.Close
210             Set rstEmploye = Nothing

215             rstModif.Fields("Date") = Replace(sNoFacture, "F-", "")
220             rstModif.Fields("Heure") = " "
225             rstModif.Fields("Type") = "M"
230             rstModif.Fields("TypeModif") = "FACTURATION"
235             rstModif.Fields("Valeur") = sPrixTotal

240             Call rstModif.Update

245             Call rstModif.Close
250             Set rstModif = Nothing
255           End If
260         End If
265       Next
270     End If

275     Exit Sub

AfficherErreur:

280     woups "frmProjSoumMec", "EnregistrerFACT", Err, Erl
End Sub

Private Sub EnregistrerProjSoum(ByVal sNoProjSoum As String)

5       On Error GoTo AfficherErreur

10      Dim rstProjSoum         As ADODB.Recordset
15      Dim rstPiece            As ADODB.Recordset
20      Dim rstEmploye          As ADODB.Recordset
25      Dim rstModif            As ADODB.Recordset
30      Dim rstOuvert           As ADODB.Recordset
35      Dim rstSection          As ADODB.Recordset
40      Dim rstPunch            As ADODB.Recordset
45      Dim iCompteur           As Integer
50      Dim itmPiece            As ListItem
55      Dim sTable              As String
60      Dim sTableModif         As String
65      Dim sTablePiece         As String
70      Dim sChamps             As String
75      Dim sSection            As String
80      Dim sExtra              As String
85      Dim bCalculExtra        As Boolean
90      Dim collExtra           As Collection
95      Dim iCompteurExtra      As Integer
100     Dim bExiste             As Boolean
105     Dim bAjoutCommande      As Boolean
110     Dim dblNbrePers         As Double
115     Dim dblJoursHebergement As Double
120     Dim dblJoursRepas       As Double
125     Dim dblHebergement1     As Double
130     Dim dblHebergement2     As Double
135     Dim dblRepas            As Double
140     Dim dblTotalHebergement As Double
           
145     Set rstEmploye = New ADODB.Recordset
         
150     Call rstEmploye.Open("SELECT noEmploye FROM GRB_Employés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)
    
155     If m_eType = TYPE_PROJET Then
160       sTable = "GRB_ProjetMec"
165       sTableModif = "GRB_Projet_Modif"
170       sTablePiece = "GRB_Projet_Pieces"
175       sChamps = "IDProjet"
180     Else
185       sTable = "GRB_SoumissionMec"
190       sTableModif = "GRB_Soumission_Modif"
195       sTablePiece = "GRB_Soumission_Pieces"
200       sChamps = "IDSoumission"
205     End If

210     Set rstProjSoum = New ADODB.Recordset
215     Set rstOuvert = New ADODB.Recordset

220     Set collExtra = New Collection
        
        'Si c'est un ajout
225     If m_bModeAjout = True Then
          'On ouvre le recordset selon le type
230       Call rstProjSoum.Open("SELECT * FROM " & sTable & " WHERE " & sChamps & " = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)

235       If m_eType = TYPE_PROJET Then
240         If rstProjSoum.EOF Then
245           bAjoutCommande = True
250         Else
255           bAjoutCommande = False
260         End If
265       Else
270         bAjoutCommande = False
275       End If

280       Call rstProjSoum.AddNew

285       If m_eType = TYPE_PROJET Then
290         rstProjSoum.Fields("LiaisonChargeable") = m_sLiaison
295       End If

300       rstProjSoum.Fields(sChamps) = sNoProjSoum

305       rstProjSoum.Fields("Creer") = ConvertDate(Date)
310       rstProjSoum.Fields("Creer_Par") = rstEmploye.Fields("noEmploye")

315       If m_eType = TYPE_PROJET Then
320         rstProjSoum.Fields("IDSoumission") = txtNoSoumission.Text
325       End If

330       Call rstOuvert.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)

335       If rstOuvert.EOF Then
340         Call rstOuvert.AddNew

345         rstOuvert.Fields("IDProjSoum") = sNoProjSoum
350         rstOuvert.Fields("NoClient") = cmbclient.ItemData(cmbclient.ListIndex)
355         rstOuvert.Fields("Description") = txtDescription.Text
360         rstOuvert.Fields("DateOuverture") = ConvertDate(Date)
365         rstOuvert.Fields("Ouvert") = True

370         If m_eType = TYPE_PROJET Then
375           rstOuvert.Fields("Type") = "P"
380         Else
385           rstOuvert.Fields("Type") = "S"
390         End If

395         Call rstOuvert.Update
400       End If

405       Call rstOuvert.Close
    
410       m_bModeAjout = False
415     Else
420       Call EnregistrerSuppression

425       Call rstProjSoum.Open("SELECT * FROM " & sTable & " WHERE " & sChamps & " = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)
        
430       Set rstModif = New ADODB.Recordset
        
435       Call rstModif.Open("SELECT * FROM " & sTableModif, g_connData, adOpenDynamic, adLockOptimistic)
        
440       Call rstModif.AddNew
       
445       rstModif.Fields("Type") = "M"
450       rstModif.Fields(sChamps) = sNoProjSoum
455       rstModif.Fields("noEmployé") = rstEmploye.Fields("noEmploye")
460       rstModif.Fields("Date") = ConvertDate(Date)
465       rstModif.Fields("Heure") = Time
470       rstModif.Fields("TypeModif") = "MODIFICATION"
       
475       Call rstModif.Update
     
480       Call rstModif.Close
485       Set rstModif = Nothing
   
490       Call rstOuvert.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)

495       If rstOuvert.Fields("NoClient") <> cmbclient.ItemData(cmbclient.ListIndex) Then
500         rstOuvert.Fields("NoClient") = cmbclient.ItemData(cmbclient.ListIndex)

505         Set rstPunch = New ADODB.Recordset

510         Call rstPunch.Open("SELECT * FROM GRB_Punch WHERE NoProjet = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)

515         If Not rstPunch.EOF Then
520           If MsgBox("Le client a été modifié, voulez-vous changer les punch de ce projet ?", vbYesNo) = vbYes Then

525             Do While Not rstPunch.EOF
530               rstPunch.Fields("NoClient") = cmbclient.ItemData(cmbclient.ListIndex)

535               Call rstPunch.Update

540               Call rstPunch.MoveNext
545             Loop
550           End If
555         End If

560         Call rstPunch.Close
565         Set rstPunch = Nothing
570       End If

575       rstOuvert.Fields("Description") = txtDescription.Text
    
580       Call rstOuvert.Update

585       Call rstOuvert.Close
590       Set rstOuvert = Nothing
        
          'Si c'est une modification, il faut effacer les pieces et remplir les nouvelles
595       Call g_connData.Execute("DELETE * FROM " & sTablePiece & " WHERE " & sChamps & " = '" & sNoProjSoum & "' AND Type = 'M'")

600       If m_eType = TYPE_PROJET Then
605         If Right$(sNoProjSoum, 2) >= 60 And Right$(sNoProjSoum, 2) <= 98 Then
610           Call g_connData.Execute("DELETE * FROM GRB_Projet_Pieces WHERE IDProjet = '" & Left$(sNoProjSoum, Len(sNoProjSoum) - 3) & "-" & rstProjSoum.Fields("LiaisonChargeable") & "' AND Type = 'M' AND (PieceExtraChargeable = True OR PieceExtraNonChargeable = True) AND Provenance = '" & Right$(txtNoProjSoum.Text, 2) & "'")
615         End If
620       End If
625     End If

630     Set rstOuvert = Nothing
       
        'Enregistrement de la soumission
635     rstProjSoum.Fields("IDClient") = cmbclient.ItemData(cmbclient.ListIndex)
640     rstProjSoum.Fields("IDContact") = cmbContact.ItemData(cmbContact.ListIndex)
645     rstProjSoum.Fields("Description") = txtDescription.Text
650     rstProjSoum.Fields("Imprevue") = m_sImprevue
655     rstProjSoum.Fields("Commission") = m_sCommission
660     rstProjSoum.Fields("Profit") = m_sProfit
    
665     If IsNumeric(txtNbreManuel.Text) Then
670       rstProjSoum.Fields("Manuel") = txtNbreManuel.Text
675     Else
680       rstProjSoum.Fields("Manuel") = "0"
685     End If
       
690     If IsNumeric(txtPrixManuel.Text) Then
695       rstProjSoum.Fields("Total_Manuel") = txtPrixManuel.Text
700     Else
705       rstProjSoum.Fields("Total_Manuel") = "0"
710     End If

715     rstProjSoum.Fields("total_Commission") = txtCommission.Text
720     rstProjSoum.Fields("Total_Profit") = txtProfit.Text
725     rstProjSoum.Fields("PrixTotal") = txtPrixTotal.Text
730     rstProjSoum.Fields("Total_piece") = txtTotalPieces.Text
735     rstProjSoum.Fields("total_imprevue") = txtImprevus.Text

740     If m_eType = TYPE_SOUMISSION Then
745       rstProjSoum.Fields("TempsDessin") = m_sTempsDessin
750       rstProjSoum.Fields("TempsCoupe") = m_sTempsCoupe
755       rstProjSoum.Fields("TempsMachinage") = m_sTempsMachinage
760       rstProjSoum.Fields("TempsSoudure") = m_sTempsSoudure
765       rstProjSoum.Fields("TempsAssemblage") = m_sTempsAssemblage
770       rstProjSoum.Fields("TempsPeinture") = m_sTempsPeinture
775       rstProjSoum.Fields("TempsTest") = m_sTempsTest
780       rstProjSoum.Fields("TempsInstallation") = m_sTempsInstallation
785       rstProjSoum.Fields("TempsFormation") = m_sTempsFormation
790       rstProjSoum.Fields("TempsGestion") = m_sTempsGestion
795       rstProjSoum.Fields("TempsShipping") = m_sTempsShipping
800     Else
805       rstProjSoum.Fields("TempsProjBarré") = m_bTempsProjLock
      
810       rstProjSoum.Fields("TempsDessinProj") = m_sTempsDessinProj
815       rstProjSoum.Fields("TempsCoupeProj") = m_sTempsCoupeProj
820       rstProjSoum.Fields("TempsMachinageProj") = m_sTempsMachinageProj
825       rstProjSoum.Fields("TempsSoudureProj") = m_sTempsSoudureProj
830       rstProjSoum.Fields("TempsAssemblageProj") = m_sTempsAssemblageProj
835       rstProjSoum.Fields("TempsPeintureProj") = m_sTempsPeintureProj
840       rstProjSoum.Fields("TempsTestProj") = m_sTempsTestProj
845       rstProjSoum.Fields("TempsInstallationProj") = m_sTempsInstallationProj
850       rstProjSoum.Fields("TempsFormationProj") = m_sTempsFormationProj
855       rstProjSoum.Fields("TempsGestionProj") = m_sTempsGestionProj
860       rstProjSoum.Fields("TempsShippingProj") = m_sTempsShippingProj

865       rstProjSoum.Fields("TempsDessinConc") = m_sTempsDessinConc
870       rstProjSoum.Fields("TempsCoupeConc") = m_sTempsCoupeConc
875       rstProjSoum.Fields("TempsMachinageConc") = m_sTempsMachinageConc
880       rstProjSoum.Fields("TempsSoudureConc") = m_sTempsSoudureConc
885       rstProjSoum.Fields("TempsAssemblageConc") = m_sTempsAssemblageConc
890       rstProjSoum.Fields("TempsPeintureConc") = m_sTempsPeintureConc
895       rstProjSoum.Fields("TempsTestConc") = m_sTempsTestConc
900       rstProjSoum.Fields("TempsInstallationConc") = m_sTempsInstallationConc
905       rstProjSoum.Fields("TempsFormationConc") = m_sTempsFormationConc
910       rstProjSoum.Fields("TempsGestionConc") = m_sTempsGestionConc
915       rstProjSoum.Fields("TempsShippingConc") = m_sTempsShippingConc
920     End If
     
925     rstProjSoum.Fields("NbrePersonne") = m_sNbrePersonne
930     rstProjSoum.Fields("TempsHebergement") = m_sTempsHebergement
935     rstProjSoum.Fields("TempsRepas") = m_sTempsRepas
940     rstProjSoum.Fields("TempsTransport") = m_sTempsTransport
945     rstProjSoum.Fields("TempsUniteMobile") = m_sTempsUniteMobile
950     rstProjSoum.Fields("PrixEmballage") = m_sPrixEmballage

955     rstProjSoum.Fields("TauxHebergement1") = m_sTauxHebergement1
960     rstProjSoum.Fields("TauxHebergement2") = m_sTauxHebergement2
965     rstProjSoum.Fields("TauxRepas") = m_sTauxRepas
970     rstProjSoum.Fields("TauxTransport") = m_sTauxTransport
975     rstProjSoum.Fields("TauxUniteMobile") = m_sTauxUniteMobile

980     rstProjSoum.Fields("TauxDessin") = m_sTauxDessin
985     rstProjSoum.Fields("TauxCoupe") = m_sTauxCoupe
990     rstProjSoum.Fields("TauxMachinage") = m_sTauxMachinage
995     rstProjSoum.Fields("TauxSoudure") = m_sTauxSoudure
1000    rstProjSoum.Fields("TauxAssemblage") = m_sTauxAssemblage
1005    rstProjSoum.Fields("TauxPeinture") = m_sTauxPeinture
1010    rstProjSoum.Fields("TauxTest") = m_sTauxTest
1015    rstProjSoum.Fields("TauxInstallation") = m_sTauxInstallation
1020    rstProjSoum.Fields("TauxFormation") = m_sTauxFormation
1025    rstProjSoum.Fields("TauxGestion") = m_sTauxGestion
1030    rstProjSoum.Fields("TauxShipping") = m_sTauxShipping

1035    rstProjSoum.Fields("CheminPhotos") = txtCheminPhotos.Text
1040    rstProjSoum.Fields("MontantForfait") = txtForfait.Text
1045    rstProjSoum.Fields("InitialeForfait") = lblForfaitInitiale.Caption

1050    If Not IsNull(rstProjSoum.Fields("NbrePersonne")) Then
1055      dblNbrePers = CDbl(rstProjSoum.Fields("NbrePersonne"))
1060    Else
1065      dblNbrePers = 0
1070    End If

1075    If Not IsNull(rstProjSoum.Fields("TempsHebergement")) Then
1080      dblJoursHebergement = CDbl(rstProjSoum.Fields("TempsHebergement"))
1085    Else
1090      dblJoursHebergement = 0
1095    End If

1100    If Not IsNull(rstProjSoum.Fields("TempsRepas")) Then
1105      dblJoursRepas = CDbl(rstProjSoum.Fields("TempsRepas"))
1110    Else
1115      dblJoursRepas = 0
1120    End If

1125    If Not IsNull(rstProjSoum.Fields("TauxHebergement1")) Then
1130      dblHebergement1 = CDbl(rstProjSoum.Fields("TauxHebergement1"))
1135    Else
1140      dblHebergement1 = 0
1145    End If

1150    If Not IsNull(rstProjSoum.Fields("TauxHebergement2")) Then
1155      dblHebergement2 = CDbl(rstProjSoum.Fields("TauxHebergement2"))
1160    Else
1165      dblHebergement2 = 0
1170    End If

1175    If Not IsNull(rstProjSoum.Fields("TauxRepas")) Then
1180      dblRepas = CDbl(rstProjSoum.Fields("TauxRepas"))
1185    Else
1190      dblRepas = 0
1195    End If

1200    rstProjSoum.Fields("TotalRepas") = dblNbrePers * dblJoursRepas * dblRepas

1205    dblTotalHebergement = 0

1210    Do While dblNbrePers > 0
1215      If dblNbrePers >= 2 Then
1220        dblTotalHebergement = dblTotalHebergement + (dblJoursHebergement * dblHebergement2)

1225        dblNbrePers = dblNbrePers - 2
1230      Else
1235        dblTotalHebergement = dblTotalHebergement + (dblJoursHebergement * dblHebergement1)

1240        dblNbrePers = dblNbrePers - 1
1245      End If
1250    Loop

1255    rstProjSoum.Fields("TotalHebergement") = dblTotalHebergement

1260    If bAjoutCommande = True Then
1265      rstProjSoum.Fields("ProchaineCommande") = 1
1270    End If

1275    If m_eType = TYPE_PROJET Then
1280      rstProjSoum.Fields("PrixRéception") = txtPrixReception.Text
1285    End If
     
        'Total des temps
1290    rstProjSoum.Fields("Total_Temps") = txtTotalTemps.Text
  
1295    Set rstPiece = New ADODB.Recordset
  
1300    If m_eType = TYPE_PROJET Then
1305      Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces", g_connData, adOpenDynamic, adLockOptimistic)
1310    Else
1315      Call rstPiece.Open("SELECT * FROM GRB_Soumission_Pieces", g_connData, adOpenDynamic, adLockOptimistic)
1320    End If
   
1325    If m_eType = TYPE_PROJET Then
1330      If g_bModificationFacturation = True Then
1335        Call EnregistrerFACT(sNoProjSoum)
1340      End If
1345    End If
  
        'Enregistrement des pièces
1350    For iCompteur = 1 To lvwSoumission.ListItems.count
1355      If lvwSoumission.ListItems(iCompteur).Tag <> vbNullString Then
1360        If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> vbNullString Then
1365          Set itmPiece = lvwSoumission.ListItems(iCompteur)
     
1370          Call rstPiece.AddNew
           
1375          If m_eType = TYPE_PROJET Then
1380            rstPiece.Fields("IDProjet") = sNoProjSoum
1385          Else
1390            rstPiece.Fields("IDSoumission") = sNoProjSoum
1395          End If
        
1400          rstPiece.Fields("Type") = "M"
          
1405          If itmPiece.Checked = True Then
1410            rstPiece.Fields("Visible") = True
1415          Else
1420            rstPiece.Fields("Visible") = False
1425          End If

1430          If m_eType = TYPE_PROJET Then
1435            rstPiece.Fields("Facturation") = itmPiece.SubItems(I_COL_SOUM_FACTURATION)

1440            If itmPiece.SubItems(I_COL_SOUM_FACTURATION) = "" Then
1445              itmPiece.SubItems(I_COL_SOUM_FACTURATION) = " "
1450            End If

1455            If itmPiece.SubItems(I_COL_SOUM_DATE_COMMANDE) = "" Then
1460              itmPiece.SubItems(I_COL_SOUM_DATE_COMMANDE) = " "
1465            End If

1470            rstPiece.Fields("NoRetour") = itmPiece.ListSubItems(I_COL_SOUM_DATE_COMMANDE).Tag
1475          End If

1480          rstPiece.Fields("IDSection") = itmPiece.Tag
1485          rstPiece.Fields("NumItem") = itmPiece.SubItems(I_COL_SOUM_PIECE)
1490          rstPiece.Fields("Qté") = Replace(itmPiece.Text, "*", "")

1495          If itmPiece.SubItems(I_COL_SOUM_PIECE) = "Texte" Or itmPiece.SubItems(I_COL_SOUM_PIECE) = "Text" Then
1500            rstPiece.Fields("DESC_EN") = ""
1505            rstPiece.Fields("DESC_FR") = itmPiece.SubItems(I_COL_SOUM_DESCR)
1510          Else
1515            If m_eLangage = ANGLAIS Then
1520              rstPiece.Fields("DESC_EN") = itmPiece.SubItems(I_COL_SOUM_DESCR)
1525              rstPiece.Fields("DESC_FR") = itmPiece.ListSubItems(I_COL_SOUM_DESCR).Tag
1530            Else
1535              rstPiece.Fields("DESC_FR") = itmPiece.SubItems(I_COL_SOUM_DESCR)
1540              rstPiece.Fields("DESC_EN") = itmPiece.ListSubItems(I_COL_SOUM_DESCR).Tag
1545            End If
1550          End If

1555          rstPiece.Fields("Manufact") = itmPiece.SubItems(I_COL_SOUM_MANUFACT)
1560          rstPiece.Fields("Prix_list") = Conversion(itmPiece.SubItems(I_COL_SOUM_PRIX_LIST), MODE_PAS_FORMAT, 4)

1565          If Trim$(itmPiece.SubItems(I_COL_SOUM_ESCOMPTE)) <> "" Then
1570            rstPiece.Fields("Escompte") = Conversion(Replace(itmPiece.SubItems(I_COL_SOUM_ESCOMPTE), "%", "") / 100, MODE_PAS_FORMAT)
1575          Else
1580            rstPiece.Fields("Escompte") = ""
1585          End If

1590          rstPiece.Fields("Prix_net") = Conversion(itmPiece.SubItems(I_COL_SOUM_PRIX_NET), MODE_PAS_FORMAT, 4)

1595          If m_eType = TYPE_PROJET Then
1600            rstPiece.Fields("DateRéception") = itmPiece.ListSubItems(I_COL_SOUM_PRIX_NET).Tag
1605          End If

1610          rstPiece.Fields("OrdreSection") = itmPiece.ListSubItems(I_COL_SOUM_MANUFACT).Tag
1615          rstPiece.Fields("NuméroLigne") = iCompteur

1620          If itmPiece.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_ORANGE Then
1625            rstPiece.Fields("Commandé") = True
1630          Else
1635            rstPiece.Fields("Commandé") = False
1640          End If

1645          If itmPiece.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_GRIS Then
1650            rstPiece.Fields("Recu") = True
1655          Else
1660            rstPiece.Fields("Recu") = False
1665          End If

1670          If itmPiece.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_ROUGE Then
1675            rstPiece.Fields("Retour") = True
1680          Else
1685            rstPiece.Fields("Retour") = False
1690          End If

1695          If itmPiece.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_VERT_FORET And itmPiece.ListSubItems(I_COL_SOUM_PIECE).Bold = True Then
1700            rstPiece.Fields("CommandeAnnulée") = True
1705          Else
1710            rstPiece.Fields("CommandeAnnulée") = False
1715          End If

1720          If itmPiece.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_BRUN Then
1725            rstPiece.Fields("MatérielInutile") = True
1730          Else
1735            rstPiece.Fields("MatérielInutile") = False
1740          End If

1745          If Trim$(itmPiece.SubItems(I_COL_SOUM_DISTRIB)) <> vbNullString Then
1750            rstPiece.Fields("IDFRS") = CInt(itmPiece.ListSubItems(I_COL_SOUM_DISTRIB).Tag)
1755          Else
1760            rstPiece.Fields("IDFRS") = 0
1765          End If

1770          rstPiece.Fields("Prix_Total") = Conversion(itmPiece.SubItems(I_COL_SOUM_TOTAL), MODE_PAS_FORMAT)
1775          rstPiece.Fields("Profit_argent") = Conversion(itmPiece.SubItems(I_COL_SOUM_PROFIT), MODE_PAS_FORMAT)

1780          If Len(itmPiece.ListSubItems(I_COL_SOUM_PIECE).Tag) <= 50 Then
1785            rstPiece.Fields("SousSection") = itmPiece.ListSubItems(I_COL_SOUM_PIECE).Tag
1790          Else
1795            rstPiece.Fields("SousSection") = Left$(itmPiece.ListSubItems(I_COL_SOUM_PIECE).Tag, 50)
1800          End If

1805          If itmPiece.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag <> vbNullString Then
1810            rstPiece.Fields("PrixOrigine") = itmPiece.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag
1815          Else
1820            rstPiece.Fields("PrixOrigine") = "0"
1825          End If

1830          If InStr(1, itmPiece.Text, "*") > 0 Then
1835            rstPiece.Fields("Quoté") = True
1840          Else
1845            rstPiece.Fields("Quoté") = False
1850          End If

1855          If m_eType = TYPE_PROJET Then
1860            If itmPiece.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_ROSE Then
1865              rstPiece.Fields("PieceExtraNonChargeable") = True
1870              rstPiece.Fields("PieceExtraChargeable") = False
1875            Else
1880              If itmPiece.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_BLEU Then
1885                rstPiece.Fields("PieceExtraChargeable") = True
1890                rstPiece.Fields("PieceExtraNonChargeable") = False
1895              Else
1900                If Left$(itmPiece.ListSubItems(I_COL_SOUM_PROFIT).Tag, 6) = "RETOUR" Or Left$(itmPiece.ListSubItems(I_COL_SOUM_PROFIT).Tag, 10) = "ANNULATION" Then
1905                  sExtra = Right$(itmPiece.ListSubItems(I_COL_SOUM_PROFIT).Tag, 2)

1910                  If sExtra >= "80" And sExtra <= "98" Then
1915                    rstPiece.Fields("PieceExtraNonChargeable") = True
1920                    rstPiece.Fields("PieceExtraChargeable") = False
1925                  Else
1930                    rstPiece.Fields("PieceExtraChargeable") = True
1935                    rstPiece.Fields("PieceExtraNonChargeable") = False
1940                  End If
1945                End If
1950              End If
1955            End If

1960            If itmPiece.SubItems(I_COL_SOUM_PROVENANCE) <> "" Then
1965              rstPiece.Fields("Provenance") = Right$(itmPiece.SubItems(I_COL_SOUM_PROVENANCE), 2)
1970            Else
1975              If Left$(itmPiece.ListSubItems(I_COL_SOUM_PROFIT).Tag, 6) = "RETOUR" Or Left$(itmPiece.ListSubItems(I_COL_SOUM_PROFIT).Tag, 10) = "ANNULATION" Then
1980                rstPiece.Fields("Provenance") = sExtra
1985              End If
1990            End If

1995            rstPiece.Fields("DateCommande") = itmPiece.SubItems(I_COL_SOUM_DATE_COMMANDE)
2000            rstPiece.Fields("DateRequise") = itmPiece.SubItems(I_COL_SOUM_DATE_REQUISE)

2005            If itmPiece.SubItems(I_COL_SOUM_DATE_REQUISE) = "" Then
2010              itmPiece.SubItems(I_COL_SOUM_DATE_REQUISE) = " "
2015            End If

2020            rstPiece.Fields("DateRetour") = itmPiece.ListSubItems(I_COL_SOUM_DATE_REQUISE).Tag

2025            rstPiece.Fields("NomCommande") = itmPiece.SubItems(I_COL_SOUM_NOM_COMMANDE)

2030            rstPiece.Fields("NoSéquentiel") = itmPiece.SubItems(I_COL_SOUM_NO_SEQUENTIEL)
2035          End If

2040          rstPiece.Fields("Commentaire") = itmPiece.SubItems(I_COL_SOUM_COMMENTAIRE)

2045          Call rstPiece.Update

2050          If m_eType = TYPE_PROJET Then
2055            If Right$(txtNoProjSoum.Text, 2) <= 98 And Right$(txtNoProjSoum.Text, 2) >= 80 Then
2060              Call AjouterPiecesExtraDansJob(itmPiece, rstProjSoum.Fields("LiaisonChargeable"))
2065            Else
2070              If Right$(txtNoProjSoum.Text, 2) <= 79 And Right$(txtNoProjSoum.Text, 2) >= 60 Then
2075                Call AjouterPiecesExtraChargeableDansJob(itmPiece, rstProjSoum.Fields("LiaisonChargeable"))
2080              Else
2085                If Left$(itmPiece.ListSubItems(I_COL_SOUM_PROFIT).Tag, 6) = "RETOUR" Then
2090                  Call AjouterInutileDansExtra(itmPiece, sExtra)

2095                  bCalculExtra = True

2100                  bExiste = False

2105                  For iCompteurExtra = 1 To collExtra.count
2110                    If collExtra(iCompteurExtra) = Right$(itmPiece.ListSubItems(I_COL_SOUM_PROFIT).Tag, 2) Then
2115                      bExiste = True

2120                      Exit For
2125                    End If
2130                  Next

2135                  If bExiste = False Then
2140                    Call collExtra.Add(Right$(itmPiece.ListSubItems(I_COL_SOUM_PROFIT).Tag, 2))
2145                  End If
2150                Else
2155                  If Left$(itmPiece.ListSubItems(I_COL_SOUM_PROFIT).Tag, 10) = "ANNULATION" Then
2160                    Call AjouterAnnulationDansExtra(itmPiece, sExtra)
  
2165                    bCalculExtra = True
  
2170                    bExiste = False
  
2175                    For iCompteurExtra = 1 To collExtra.count
2180                      If collExtra(iCompteurExtra) = Right$(itmPiece.ListSubItems(I_COL_SOUM_PROFIT).Tag, 2) Then
2185                        bExiste = True
  
2190                        Exit For
2195                      End If
2200                    Next
  
2205                    If bExiste = False Then
2210                      Call collExtra.Add(Right$(itmPiece.ListSubItems(I_COL_SOUM_PROFIT).Tag, 2))
2215                    End If
2220                  End If
2225                End If
2230              End If
2235            End If
2240          End If
2245        End If
2250      End If
2255    Next

2260    If m_eType = TYPE_PROJET Then
2265      If Right$(txtNoProjSoum.Text, 2) <= 98 And Right$(txtNoProjSoum.Text, 2) >= 60 Then
2270        Call CalculerTotalRecordset(Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & rstProjSoum.Fields("LiaisonChargeable"))
2275      End If
2280    End If

2285    If bCalculExtra = True Then
2290      For iCompteurExtra = 1 To collExtra.count
2295        Call CalculerTotalRecordset(Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & collExtra(iCompteurExtra))
2300      Next
2305    End If

2310    Call rstProjSoum.Update

2315    Call rstProjSoum.Close
2320    Set rstProjSoum = Nothing
  
2325    Call rstPiece.Close
2330    Set rstPiece = Nothing

2335    If m_eType = TYPE_SOUMISSION Then
2340      Call AjouterSoumissionAuCumulatif
2345    Else
2350      Call AjouterProjetAuCumulatif
2355    End If

2360    Exit Sub

AfficherErreur:

2365    woups "frmProjSoumMec", "EnregistrerProjSoum", Err, Erl, sNoProjSoum)

  'Si un erreur se produit dans l'enregistrement des pièces, il faut avertir
  'l'utilisateur de quelle pièce il s'agit et continuer avec un Resume Next
2370    If Erl >= 1355 And Erl <= 2250 Then
2375      Set rstSection = New ADODB.Recordset
  
2380      If m_eLangage = ANGLAIS Then
2385        Call rstSection.Open("SELECT NomSectionEN FROM GRB_SoumProjSectionMec", g_connData, adOpenDynamic, adLockOptimistic)

2390        If Not rstSection.EOF Then
2395          sSection = rstSection.Fields("NomSectionEN")
2400        Else
2405          sSection = itmPiece.Tag
2410        End If
2415      Else
2420        Call rstSection.Open("SELECT NomSectionFR FROM GRB_SoumProjSectionMec", g_connData, adOpenDynamic, adLockOptimistic)

2425        If Not rstSection.EOF Then
2430          sSection = rstSection.Fields("NomSectionFR")
2435        Else
2440          sSection = itmPiece.Tag
2445        End If
2450      End If

2455      Call rstSection.Close
2460      Set rstSection = Nothing

2465      Call MsgBox("La pièce " & itmPiece.SubItems(I_COL_SOUM_PIECE) & " de la section " & sSection & " risque de contenir des erreurs." & vbNewLine & _
               "Il se peut qu'elle ne soit plus présente dans la liste.")
2470    End If

2475    Resume Next
End Sub

Private Sub AjouterPiecesExtraChargeableDansJob(ByVal itmSource As ListItem, ByVal sLiaison As String)

5       On Error GoTo AfficherErreur

10      Dim rstPiece   As ADODB.Recordset
15      Dim rstProjet  As ADODB.Recordset
20      Dim rstSection As ADODB.Recordset
25      Dim iCompteur  As Integer
30      Dim sSection   As String
35      Dim bSkip      As Boolean
        
40      Set rstProjet = New ADODB.Recordset
        
45      Call rstProjet.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sLiaison & "'", g_connData, adOpenDynamic, adLockOptimistic)

        'Si le projet existe
55      If Not rstProjet.EOF Then
          'Ouverture du recordset sur le projet original
60        Set rstPiece = New ADODB.Recordset
          
65        Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sLiaison & "' AND IDSection = " & itmSource.Tag & " AND SousSection = '" & Replace(itmSource.ListSubItems(I_COL_SOUM_PIECE).Tag, "'", "''") & "' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)

70        If Not rstPiece.EOF Then
75          Call rstPiece.MoveLast

80          iCompteur = rstPiece.Fields("NuméroLigne") + 1
85        Else
90          iCompteur = 1
95        End If

100       Call rstPiece.AddNew

105       rstPiece.Fields("IDProjet") = Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sLiaison
       
110       rstPiece.Fields("Type") = "M"
       
115       If itmSource.Checked = True Then
120         rstPiece.Fields("Visible") = True
125       Else
130         rstPiece.Fields("Visible") = False
135       End If

140       rstPiece.Fields("Facturation") = itmSource.SubItems(I_COL_SOUM_FACTURATION)
              
145       rstPiece.Fields("IDSection") = itmSource.Tag
150       rstPiece.Fields("NumItem") = Trim$(itmSource.SubItems(I_COL_SOUM_PIECE))
155       rstPiece.Fields("Qté") = Replace(itmSource.Text, "*", vbNullString)
160       rstPiece.Fields("Desc_FR") = Trim$(itmSource.SubItems(I_COL_SOUM_DESCR))
165       rstPiece.Fields("Desc_EN") = Trim$(itmSource.ListSubItems(I_COL_SOUM_DESCR).Tag)
170       rstPiece.Fields("Manufact") = Trim$(itmSource.SubItems(I_COL_SOUM_MANUFACT))
175       rstPiece.Fields("Prix_list") = Conversion(itmSource.SubItems(I_COL_SOUM_PRIX_LIST), MODE_PAS_FORMAT, 4)
180       rstPiece.Fields("Escompte") = Conversion(itmSource.SubItems(I_COL_SOUM_ESCOMPTE), MODE_PAS_FORMAT)
185       rstPiece.Fields("Prix_net") = Conversion(itmSource.SubItems(I_COL_SOUM_PRIX_NET), MODE_PAS_FORMAT, 4)
190       rstPiece.Fields("OrdreSection") = itmSource.ListSubItems(I_COL_SOUM_MANUFACT).Tag
195       rstPiece.Fields("NuméroLigne") = iCompteur
      
200       If itmSource.SubItems(I_COL_SOUM_DISTRIB) <> vbNullString Then
205         rstPiece.Fields("IDFRS") = itmSource.ListSubItems(I_COL_SOUM_DISTRIB).Tag
210       End If
        
215       rstPiece.Fields("Prix_Total") = Conversion(itmSource.SubItems(I_COL_SOUM_TOTAL), MODE_PAS_FORMAT)
220       rstPiece.Fields("Profit_argent") = Conversion(itmSource.SubItems(I_COL_SOUM_PROFIT), MODE_PAS_FORMAT)
225       rstPiece.Fields("SousSection") = itmSource.ListSubItems(I_COL_SOUM_PIECE).Tag
            
230       If itmSource.SubItems(I_COL_SOUM_PRIX_LIST) <> vbNullString Then
235         If itmSource.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag <> vbNullString Then
240           rstPiece.Fields("PrixOrigine") = Replace(Round(CDbl(Replace(itmSource.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag, ".", ",")), 4), ".", ",")
245         Else
250           rstPiece.Fields("PrixOrigine") = "0"
255         End If
260       Else
265         rstPiece.Fields("PrixOrigine") = "0"
270       End If
    
275       If InStr(1, itmSource.Text, "*") > 0 Then
280         rstPiece.Fields("Quoté") = True
285       Else
290         rstPiece.Fields("Quoté") = False
295       End If

300       rstPiece.Fields("PieceExtraChargeable") = True
305       rstPiece.Fields("Provenance") = Right$(txtNoProjSoum.Text, 2)

310       Call rstPiece.Update

315       Call rstPiece.Close

320       rstPiece.CursorLocation = adUseServer

325       Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sLiaison & "' AND NuméroLigne >= " & iCompteur & " ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)

          'Tant qu'il y a des enregistrements dans le recordset
330       Do While Not rstPiece.EOF
335         If rstPiece.Fields("PieceExtraChargeable") = True And rstPiece.Fields("Qté") = itmSource.Text And rstPiece.Fields("NumItem") = itmSource.SubItems(I_COL_SOUM_PIECE) And bSkip = False Then
340           bSkip = True
345         Else
350           rstPiece.Fields("NuméroLigne") = rstPiece.Fields("NuméroLigne") + 1

355           Call rstPiece.Update
360         End If

365         Call rstPiece.MoveNext
370       Loop

375       Call rstPiece.Close
380       Set rstPiece = Nothing
385     End If

390     Call rstProjet.Close
395     Set rstProjet = Nothing

400     Exit Sub

AfficherErreur:

405     woups "frmProjSoumMec", "AjouterPiecesExtraDansJob", Err, Erl

  Set rstSection = New ADODB.Recordset

  If m_eLangage = ANGLAIS Then
    Call rstSection.Open("SELECT NomSectionEN FROM GRB_SoumProjSectionMec", g_connData, adOpenDynamic, adLockOptimistic)
      
    If Not rstSection.EOF Then
      sSection = rstSection.Fields("NomSectionEN")
    Else
      sSection = itmSource.Tag
    End If
  Else
    Call rstSection.Open("SELECT NomSectionFR FROM GRB_SoumProjSectionMec", g_connData, adOpenDynamic, adLockOptimistic)
      
    If Not rstSection.EOF Then
      sSection = rstSection.Fields("NomSectionFR")
    Else
      sSection = itmSource.Tag
    End If
  End If
     
  Call rstSection.Close
  Set rstSection = Nothing
  
  Call MsgBox("La pièce " & itmSource.SubItems(I_COL_SOUM_PIECE) & " de la section " & sSection & " risque de contenir des erreurs dans le projet " & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sLiaison & "." & vbNewLine & _
              "Il se peut qu'elle ne soit pas ajoutée dans le projet.")

  Resume Next
End Sub

Private Sub AjouterPiecesExtraDansJob(ByVal itmSource As ListItem, ByVal sLiaison As String)

5       On Error GoTo AfficherErreur

10      Dim rstPiece   As ADODB.Recordset
15      Dim rstProjet  As ADODB.Recordset
20      Dim rstSection As ADODB.Recordset
25      Dim iCompteur  As Integer
30      Dim sSection   As String
35      Dim bSkip      As Boolean

40      Set rstProjet = New ADODB.Recordset

45      Call rstProjet.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sLiaison & "'", g_connData, adOpenDynamic, adLockOptimistic)

        'Si le projet existe
50      If Not rstProjet.EOF Then
          'Ouverture du recordset
55        Set rstPiece = New ADODB.Recordset

60        Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sLiaison & "' AND IDSection = " & itmSource.Tag & " AND SousSection = '" & Replace(itmSource.ListSubItems(I_COL_SOUM_PIECE).Tag, "'", "''") & "' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)

65        If Not rstPiece.EOF Then
70          Call rstPiece.MoveLast

75          iCompteur = rstPiece.Fields("NuméroLigne") + 1
80        Else
85          iCompteur = 1
90        End If

95        Call rstPiece.AddNew

100       rstPiece.Fields("IDProjet") = Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sLiaison

105       rstPiece.Fields("Type") = "M"

110       If itmSource.Checked = True Then
115         rstPiece.Fields("Visible") = True
120       Else
125         rstPiece.Fields("Visible") = False
130       End If

135       rstPiece.Fields("Facturation") = itmSource.SubItems(I_COL_SOUM_FACTURATION)

140       rstPiece.Fields("IDSection") = itmSource.Tag
145       rstPiece.Fields("NumItem") = Trim$(itmSource.SubItems(I_COL_SOUM_PIECE))
150       rstPiece.Fields("Qté") = Replace(itmSource.Text, "*", "")
155       rstPiece.Fields("Desc_FR") = Trim$(itmSource.SubItems(I_COL_SOUM_DESCR))
160       rstPiece.Fields("Desc_EN") = Trim$(itmSource.ListSubItems(I_COL_SOUM_DESCR).Tag)
165       rstPiece.Fields("Manufact") = Trim$(itmSource.SubItems(I_COL_SOUM_MANUFACT))
170       rstPiece.Fields("Prix_List") = Conversion(itmSource.SubItems(I_COL_SOUM_PRIX_LIST), MODE_PAS_FORMAT, 4)
175       rstPiece.Fields("Escompte") = Conversion(itmSource.SubItems(I_COL_SOUM_ESCOMPTE), MODE_PAS_FORMAT)
180       rstPiece.Fields("Prix_net") = Conversion(itmSource.SubItems(I_COL_SOUM_PRIX_NET), MODE_PAS_FORMAT, 4)
185       rstPiece.Fields("OrdreSection") = itmSource.ListSubItems(I_COL_SOUM_MANUFACT).Tag
190       rstPiece.Fields("NuméroLigne") = iCompteur

195       If itmSource.SubItems(I_COL_SOUM_DISTRIB) <> vbNullString Then
200         rstPiece.Fields("IDFRS") = CInt(itmSource.ListSubItems(I_COL_SOUM_DISTRIB).Tag)
205       End If

210       rstPiece.Fields("Prix_Total") = Conversion(itmSource.SubItems(I_COL_SOUM_TOTAL), MODE_PAS_FORMAT)
215       rstPiece.Fields("Profit_argent") = Conversion(itmSource.SubItems(I_COL_SOUM_PROFIT), MODE_PAS_FORMAT)

220       If Len(rstPiece.Fields("SousSection")) >= 255 Then
225         rstPiece.Fields("SousSection") = itmSource.ListSubItems(I_COL_SOUM_PIECE).Tag
230       Else
235         rstPiece.Fields("SousSection") = Left$(itmSource.ListSubItems(I_COL_SOUM_PIECE).Tag, 255)
240       End If

245       If itmSource.SubItems(I_COL_SOUM_PRIX_LIST) <> vbNullString Then
250         If itmSource.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag <> vbNullString Then
255           rstPiece.Fields("PrixOrigine") = Replace(Round(CDbl(Replace(itmSource.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag, ".", ",")), 2), ".", ",")
260         Else
265           rstPiece.Fields("PrixOrigine") = "0"
270         End If
275       Else
280         rstPiece.Fields("PrixOrigine") = "0"
285       End If

290       If InStr(1, itmSource.Text, "*") > 0 Then
295         rstPiece.Fields("Quoté") = True
300       Else
305         rstPiece.Fields("Quoté") = False
310       End If

315       rstPiece.Fields("PieceExtraNonChargeable") = True
320       rstPiece.Fields("Provenance") = Right$(txtNoProjSoum.Text, 2)

325       Call rstPiece.Update

330       Call rstPiece.Close

335       Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sLiaison & "' AND NuméroLigne >= " & iCompteur & " ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)

          'Tant qu'il y a des enregistrements dans le recordset
340       Do While Not rstPiece.EOF
345         If rstPiece.Fields("PieceExtraNonChargeable") = True And rstPiece.Fields("Qté") = itmSource.Text And rstPiece.Fields("NumItem") = itmSource.SubItems(I_COL_SOUM_PIECE) And bSkip = False Then
350           bSkip = True
355         Else
360           rstPiece.Fields("NuméroLigne") = rstPiece.Fields("NuméroLigne") + 1

365           Call rstPiece.Update
370         End If

375         Call rstPiece.MoveNext
380       Loop

385       Call rstPiece.Close
390       Set rstPiece = Nothing
395     End If

400     Call rstProjet.Close
405     Set rstProjet = Nothing

410     Exit Sub

AfficherErreur:

415     woups "frmProjSoumMec", "AjouterPiecesExtraDansJob", Err, Erl

  Set rstSection = New ADODB.Recordset

  If m_eLangage = ANGLAIS Then
    Call rstSection.Open("SELECT NomSectionEN FROM GRB_SoumProjSectionMec", g_connData, adOpenDynamic, adLockOptimistic)

    If Not rstSection.EOF Then
      sSection = rstSection.Fields("NomSectionEN")
    Else
      sSection = itmSource.Tag
    End If
  Else
    Call rstSection.Open("SELECT NomSectionFR FROM GRB_SoumProjSectionMec", g_connData, adOpenDynamic, adLockOptimistic)

    If Not rstSection.EOF Then
      sSection = rstSection.Fields("NomSectionFR")
    Else
      sSection = itmSource.Tag
    End If
  End If

  Call rstSection.Close
  Set rstSection = Nothing

  Call MsgBox("La pièce " & itmSource.SubItems(I_COL_SOUM_PIECE) & " de la section " & sSection & " risque de contenir des erreurs dans le projet " & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sLiaison & "." & vbNewLine & _
              "Il se peut qu'elle ne soit pas ajoutée dans le projet.")

  Resume Next
End Sub

Private Sub AjouterInutileDansExtra(ByVal itmSource As ListItem, ByVal sExtra As String)
  
5       On Error GoTo AfficherErreur

10      Dim rstPiece   As ADODB.Recordset
15      Dim rstProjet  As ADODB.Recordset
20      Dim rstSection As ADODB.Recordset
25      Dim iCompteur  As Integer
30      Dim sSection   As String
35      Dim bSkip      As Boolean

40      Set rstProjet = New ADODB.Recordset

45      Call rstProjet.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sExtra & "'", g_connData, adOpenDynamic, adLockOptimistic)

        'Si le projet existe
50      If Not rstProjet.EOF Then
          'Ouverture du recordset sur le projet original
55        Set rstPiece = New ADODB.Recordset
          
60        Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sExtra & "' AND IDSection = " & itmSource.Tag & " AND SousSection = '" & Replace(itmSource.ListSubItems(I_COL_SOUM_PIECE).Tag, "'", "''") & "' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)

65        If Not rstPiece.EOF Then
70          Call rstPiece.MoveLast

75          iCompteur = rstPiece.Fields("NuméroLigne") + 1
80        Else
85          iCompteur = 1
90        End If

95        Call rstPiece.AddNew

100       rstPiece.Fields("IDProjet") = Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sExtra
       
105       rstPiece.Fields("Type") = "M"
       
110       If itmSource.Checked = True Then
115         rstPiece.Fields("Visible") = True
120       Else
125         rstPiece.Fields("Visible") = False
130       End If
               
135       rstPiece.Fields("IDSection") = itmSource.Tag
140       rstPiece.Fields("NumItem") = Trim$(itmSource.SubItems(I_COL_SOUM_PIECE))
145       rstPiece.Fields("Qté") = Replace(itmSource.Text, "*", vbNullString)

150       If m_eLangage = ANGLAIS Then
155         rstPiece.Fields("DESC_EN") = Trim$(itmSource.SubItems(I_COL_SOUM_DESCR))
160         rstPiece.Fields("DESC_FR") = Trim$(itmSource.ListSubItems(I_COL_SOUM_DESCR).Tag)
165       Else
170         rstPiece.Fields("DESC_FR") = Trim$(itmSource.SubItems(I_COL_SOUM_DESCR))
175         rstPiece.Fields("DESC_EN") = Trim$(itmSource.ListSubItems(I_COL_SOUM_DESCR).Tag)
180       End If

185       rstPiece.Fields("Manufact") = Trim$(itmSource.SubItems(I_COL_SOUM_MANUFACT))
190       rstPiece.Fields("Prix_list") = Conversion(itmSource.SubItems(I_COL_SOUM_PRIX_LIST), MODE_PAS_FORMAT, 4)
195       rstPiece.Fields("Escompte") = Conversion(itmSource.SubItems(I_COL_SOUM_ESCOMPTE), MODE_PAS_FORMAT)
200       rstPiece.Fields("Prix_net") = Conversion(itmSource.SubItems(I_COL_SOUM_PRIX_NET), MODE_PAS_FORMAT, 4)

205       rstPiece.Fields("OrdreSection") = itmSource.ListSubItems(I_COL_SOUM_MANUFACT).Tag
210       rstPiece.Fields("NuméroLigne") = iCompteur
            
215       If itmSource.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_NOIR Then
220         rstPiece.Fields("MatérielInutile") = False
225       Else
230         rstPiece.Fields("MatérielInutile") = True
235       End If
     
240       If itmSource.SubItems(I_COL_SOUM_DISTRIB) <> vbNullString Then
245         rstPiece.Fields("IDFRS") = itmSource.ListSubItems(I_COL_SOUM_DISTRIB).Tag
250       End If
     
255       rstPiece.Fields("Prix_Total") = Conversion(itmSource.SubItems(I_COL_SOUM_TOTAL), MODE_PAS_FORMAT)
260       rstPiece.Fields("Profit_argent") = Conversion(itmSource.SubItems(I_COL_SOUM_PROFIT), MODE_PAS_FORMAT)
265       rstPiece.Fields("SousSection") = itmSource.ListSubItems(I_COL_SOUM_PIECE).Tag
            
270       If itmSource.SubItems(I_COL_SOUM_PRIX_LIST) <> vbNullString Then
275         If itmSource.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag <> vbNullString Then
280           rstPiece.Fields("PrixOrigine") = Replace(Round(CDbl(Replace(itmSource.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag, ".", ",")), 4), ".", ",")
285         Else
290           rstPiece.Fields("PrixOrigine") = "0"
295         End If
300       Else
305         rstPiece.Fields("PrixOrigine") = "0"
310       End If
  
315       If InStr(1, itmSource.Text, "*") > 0 Then
320         rstPiece.Fields("Quoté") = True
325       Else
330         rstPiece.Fields("Quoté") = False
335       End If

340       rstPiece.Fields("DateRetour") = itmSource.ListSubItems(I_COL_SOUM_DATE_REQUISE).Tag

345       rstPiece.Fields("Commentaire") = itmSource.SubItems(I_COL_SOUM_COMMENTAIRE)

350       Call rstPiece.Update

355       Call rstPiece.Close

360       rstPiece.CursorLocation = adUseServer

365       Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sExtra & "' AND NuméroLigne >= " & iCompteur & " ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)

          'Tant qu'il y a des enregistrements dans le recordset
370       Do While Not rstPiece.EOF
375         If itmSource.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_BRUN Then
380           If rstPiece.Fields("MatérielInutile") = True And rstPiece.Fields("Qté") = itmSource.Text And rstPiece.Fields("NumItem") = itmSource.SubItems(I_COL_SOUM_PIECE) And bSkip = False Then
385             bSkip = True
390           Else
395             rstPiece.Fields("NuméroLigne") = rstPiece.Fields("NuméroLigne") + 1

400             Call rstPiece.Update
405           End If
410         Else
415           If rstPiece.Fields("MatérielInutile") = False And rstPiece.Fields("Qté") = itmSource.Text And rstPiece.Fields("NumItem") = itmSource.SubItems(I_COL_SOUM_PIECE) And bSkip = False Then
420             bSkip = True
425           Else
430             rstPiece.Fields("NuméroLigne") = rstPiece.Fields("NuméroLigne") + 1

435             Call rstPiece.Update
440           End If
445         End If

450         Call rstPiece.MoveNext
455       Loop

460       Call rstPiece.Close
465       Set rstPiece = Nothing
470     End If

475     Call rstProjet.Close
480     Set rstProjet = Nothing

485     Exit Sub

AfficherErreur:

490     woups "frmProjSoumMec", "AjouterInutileDansExtra", Err, Erl

  Set rstSection = New ADODB.Recordset

  If m_eLangage = ANGLAIS Then
    Call rstSection.Open("SELECT NomSectionEN FROM GRB_SoumProjSectionElec", g_connData, adOpenDynamic, adLockOptimistic)
      
    If Not rstSection.EOF Then
      sSection = rstSection.Fields("NomSectionEN")
    Else
      sSection = itmSource.Tag
    End If
  Else
    Call rstSection.Open("SELECT NomSectionFR FROM GRB_SoumProjSectionElec", g_connData, adOpenDynamic, adLockOptimistic)
      
    If Not rstSection.EOF Then
      sSection = rstSection.Fields("NomSectionFR")
    Else
      sSection = itmSource.Tag
    End If
  End If
     
  Call rstSection.Close
  Set rstSection = Nothing
  
  Call MsgBox("La pièce " & itmSource.SubItems(I_COL_SOUM_PIECE) & " de la section " & sSection & " risque de contenir des erreurs dans le projet " & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & Right$("0" & Right$(txtNoProjSoum.Text, 2) + 80, 2) & "." & vbNewLine & _
              "Il se peut qu'elle ne soit pas ajoutée dans le projet.")
 
  Resume Next
End Sub

Private Sub AjouterAnnulationDansExtra(ByVal itmSource As ListItem, ByVal sExtra As String)
  
5       On Error GoTo AfficherErreur

10      Dim rstPiece   As ADODB.Recordset
15      Dim rstProjet  As ADODB.Recordset
20      Dim rstSection As ADODB.Recordset
25      Dim iCompteur  As Integer
30      Dim sSection   As String
35      Dim bSkip      As Boolean
        
40      Set rstProjet = New ADODB.Recordset
        
45      Call rstProjet.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sExtra & "'", g_connData, adOpenDynamic, adLockOptimistic)

        'Si le projet existe
50      If Not rstProjet.EOF Then
          'Ouverture du recordset sur le projet original
55        Set rstPiece = New ADODB.Recordset
          
60        Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sExtra & "' AND IDSection = " & itmSource.Tag & " AND SousSection = '" & Replace(itmSource.ListSubItems(I_COL_SOUM_PIECE).Tag, "'", "''") & "' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)

65        If Not rstPiece.EOF Then
70          Call rstPiece.MoveLast

75          iCompteur = rstPiece.Fields("NuméroLigne") + 1
80        Else
85          iCompteur = 1
90        End If

95        Call rstPiece.AddNew

100       rstPiece.Fields("IDProjet") = Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sExtra
       
105       rstPiece.Fields("Type") = "M"
       
110       If itmSource.Checked = True Then
115         rstPiece.Fields("Visible") = True
120       Else
125         rstPiece.Fields("Visible") = False
130       End If
               
135       rstPiece.Fields("IDSection") = itmSource.Tag
140       rstPiece.Fields("NumItem") = Trim$(itmSource.SubItems(I_COL_SOUM_PIECE))
145       rstPiece.Fields("Qté") = Replace(itmSource.Text, "*", vbNullString)

150       If m_eLangage = ANGLAIS Then
155         rstPiece.Fields("DESC_EN") = Trim$(itmSource.SubItems(I_COL_SOUM_DESCR))
160         rstPiece.Fields("DESC_FR") = Trim$(itmSource.ListSubItems(I_COL_SOUM_DESCR).Tag)
165       Else
170         rstPiece.Fields("DESC_FR") = Trim$(itmSource.SubItems(I_COL_SOUM_DESCR))
175         rstPiece.Fields("DESC_EN") = Trim$(itmSource.ListSubItems(I_COL_SOUM_DESCR).Tag)
180       End If

185       rstPiece.Fields("Manufact") = Trim$(itmSource.SubItems(I_COL_SOUM_MANUFACT))
190       rstPiece.Fields("Prix_list") = Conversion(itmSource.SubItems(I_COL_SOUM_PRIX_LIST), MODE_PAS_FORMAT, 4)
195       rstPiece.Fields("Escompte") = Conversion(itmSource.SubItems(I_COL_SOUM_ESCOMPTE), MODE_PAS_FORMAT)
200       rstPiece.Fields("Prix_net") = Conversion(itmSource.SubItems(I_COL_SOUM_PRIX_NET), MODE_PAS_FORMAT, 4)

205       rstPiece.Fields("OrdreSection") = itmSource.ListSubItems(I_COL_SOUM_MANUFACT).Tag
210       rstPiece.Fields("NuméroLigne") = iCompteur
            
215       rstPiece.Fields("CommandeAnnulée") = True
     
220       If itmSource.SubItems(I_COL_SOUM_DISTRIB) <> vbNullString Then
225         rstPiece.Fields("IDFRS") = itmSource.ListSubItems(I_COL_SOUM_DISTRIB).Tag
230       End If
     
235       rstPiece.Fields("Prix_Total") = Conversion(itmSource.SubItems(I_COL_SOUM_TOTAL), MODE_PAS_FORMAT)
240       rstPiece.Fields("Profit_argent") = Conversion(itmSource.SubItems(I_COL_SOUM_PROFIT), MODE_PAS_FORMAT)
245       rstPiece.Fields("SousSection") = itmSource.ListSubItems(I_COL_SOUM_PIECE).Tag
            
250       If itmSource.SubItems(I_COL_SOUM_PRIX_LIST) <> vbNullString Then
255         If itmSource.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag <> vbNullString Then
260           rstPiece.Fields("PrixOrigine") = Replace(Round(CDbl(Replace(itmSource.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag, ".", ",")), 2), ".", ",")
265         Else
270           rstPiece.Fields("PrixOrigine") = "0"
275         End If
280       Else
285         rstPiece.Fields("PrixOrigine") = "0"
290       End If
  
295       If InStr(1, itmSource.Text, "*") > 0 Then
300         rstPiece.Fields("Quoté") = True
305       Else
310         rstPiece.Fields("Quoté") = False
315       End If

320       rstPiece.Fields("Commentaire") = itmSource.SubItems(I_COL_SOUM_COMMENTAIRE)

325       Call rstPiece.Update

330       Call rstPiece.Close

335       rstPiece.CursorLocation = adUseServer

340       Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sExtra & "' AND NuméroLigne >= " & iCompteur & " ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)

          'Tant qu'il y a des enregistrements dans le recordset
345       Do While Not rstPiece.EOF
350         If rstPiece.Fields("CommandeAnnulée") = True And rstPiece.Fields("Qté") = itmSource.Text And rstPiece.Fields("NumItem") = itmSource.SubItems(I_COL_SOUM_PIECE) And bSkip = False Then
355           bSkip = True
360         Else
365           rstPiece.Fields("NuméroLigne") = rstPiece.Fields("NuméroLigne") + 1

370           Call rstPiece.Update
375         End If

380         Call rstPiece.MoveNext
385       Loop

390       Call rstPiece.Close
395       Set rstPiece = Nothing
400     End If

405     Call rstProjet.Close
410     Set rstProjet = Nothing

415     Exit Sub

AfficherErreur:

420     woups "frmProjSoumMec", "AjouterAnnulationDansExtra", Err, Erl

  Set rstSection = New ADODB.Recordset

  If m_eLangage = ANGLAIS Then
    Call rstSection.Open("SELECT NomSectionEN FROM GRB_SoumProjSectionElec", g_connData, adOpenDynamic, adLockOptimistic)
      
    If Not rstSection.EOF Then
      sSection = rstSection.Fields("NomSectionEN")
    Else
      sSection = itmSource.Tag
    End If
  Else
    Call rstSection.Open("SELECT NomSectionFR FROM GRB_SoumProjSectionElec", g_connData, adOpenDynamic, adLockOptimistic)
      
    If Not rstSection.EOF Then
      sSection = rstSection.Fields("NomSectionFR")
    Else
      sSection = itmSource.Tag
    End If
  End If
     
  Call rstSection.Close
  Set rstSection = Nothing
  
  Call MsgBox("La pièce " & itmSource.SubItems(I_COL_SOUM_PIECE) & " de la section " & sSection & " risque de contenir des erreurs dans le projet " & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & Right$("0" & Right$(txtNoProjSoum.Text, 2) + 80, 2) & "." & vbNewLine & _
              "Il se peut qu'elle ne soit pas ajoutée dans le projet.")
 
  Resume Next
End Sub

Private Sub Cmdfermer_Click()

5       On Error GoTo AfficherErreur

        'Fermeture de la fenêtre
10      m_bResize = False
  
15      Call Unload(Me)

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumMec", "cmdFermer_Click", Err, Erl
End Sub

Private Sub RemplirListViewModifications()

5       On Error GoTo AfficherErreur

        'Rempli le lvwHistorique
10      Dim rstProjSoum As ADODB.Recordset
15      Dim rstEmploye  As ADODB.Recordset
20      Dim rstCreation As ADODB.Recordset
25      Dim sChamps     As String
30      Dim sTable      As String
35      Dim sTableCreer As String
40      Dim itmModif    As ListItem
    
        'Il faut le vider avant de le remplir
45      Call lvwHistorique.ListItems.Clear
          
50      If m_eType = TYPE_PROJET Then
55        sChamps = "IDProjet"
60        sTable = "GRB_Projet_Modif"
65        sTableCreer = "GRB_ProjetMec"
70      Else
75        sChamps = "IDSoumission"
80        sTable = "GRB_Soumission_Modif"
85        sTableCreer = "GRB_SoumissionMec"
90      End If
  
        'Ouverture du recordset selon le type
95      Set rstEmploye = New ADODB.Recordset
100     Set rstCreation = New ADODB.Recordset
        
105     Call rstCreation.Open("SELECT creer , creer_par FROM " & sTableCreer & " WHERE " & sChamps & " = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
    
        'Ajout de la section "Création"
110     Set itmModif = lvwHistorique.ListItems.Add
      
115     itmModif.Text = "Création"
    
120     itmModif.Bold = True
      
        'Ajout du nom de celui qui l'a créé
125     Set itmModif = lvwHistorique.ListItems.Add
    
130     Call rstEmploye.Open("SELECT Employe FROM GRB_Employés WHERE noEmploye = " & rstCreation.Fields("creer_par"), g_connData, adOpenDynamic, adLockOptimistic)
    
135     itmModif.Text = rstEmploye.Fields("Employe")
    
140     Call rstEmploye.Close
    
        'Date
145     itmModif.SubItems(I_COL_MODIF_DATE) = rstCreation.Fields("Creer")
   
150     itmModif.SubItems(I_COL_MODIF_HEURE) = vbNullString
   
155     Call rstCreation.Close
160     Set rstCreation = Nothing
    
165     Set rstProjSoum = New ADODB.Recordset
    
170     Call rstProjSoum.Open("SELECT * FROM " & sTable & " WHERE " & sChamps & " = '" & txtNoProjSoum.Text & "' AND Type = 'M' And TypeModif = 'MODIFICATION' ORDER BY Date, Heure", g_connData, adOpenDynamic, adLockOptimistic)

175     If Not rstProjSoum.EOF Then
          'Ajout de la section "Modifications"
180       Set itmModif = lvwHistorique.ListItems.Add
      
185       itmModif.Text = "Modifications"
    
190       itmModif.Bold = True
      
195       Do While Not rstProjSoum.EOF
            'Ajout des modifications
200         Set itmModif = lvwHistorique.ListItems.Add
        
205         Call rstEmploye.Open("SELECT Employe FROM GRB_Employés WHERE noEmploye = " & rstProjSoum.Fields("NoEmployé"), g_connData, adOpenDynamic, adLockOptimistic)
        
            'Employé
210         itmModif.Text = rstEmploye.Fields("Employe")
        
215         Call rstEmploye.Close
        
            'Date
220         itmModif.SubItems(I_COL_MODIF_DATE) = rstProjSoum.Fields("Date")
        
            'Heure
225         itmModif.SubItems(I_COL_MODIF_HEURE) = rstProjSoum.Fields("Heure")
        
230         Call rstProjSoum.MoveNext
235       Loop
240     End If
    
245     Call rstProjSoum.Close

250     Call rstProjSoum.Open("SELECT * FROM " & sTable & " WHERE " & sChamps & " = '" & txtNoProjSoum.Text & "' AND Type = 'M' AND TypeModif = 'RECEPTION' ORDER BY Date, Heure", g_connData, adOpenDynamic, adLockOptimistic)
               
255     If Not rstProjSoum.EOF Then
          'Ajout de la section "Réception"
260       Set itmModif = lvwHistorique.ListItems.Add
      
265       itmModif.Text = "Réception"
    
270       itmModif.Bold = True
      
275       Do While Not rstProjSoum.EOF
            'Ajout des modifications
280         Set itmModif = lvwHistorique.ListItems.Add
        
285         Call rstEmploye.Open("SELECT Employe FROM GRB_Employés WHERE noEmploye = " & rstProjSoum.Fields("NoEmployé"), g_connData, adOpenDynamic, adLockOptimistic)
        
            'Employé
290         itmModif.Text = rstEmploye.Fields("Employe")
        
295         Call rstEmploye.Close
        
            'Date
300         itmModif.SubItems(I_COL_MODIF_DATE) = rstProjSoum.Fields("Date")
        
            'Heure
305         itmModif.SubItems(I_COL_MODIF_HEURE) = rstProjSoum.Fields("Heure")
        
310         Call rstProjSoum.MoveNext
315       Loop
320     End If
    
325     Call rstProjSoum.Close

330     Call rstProjSoum.Open("SELECT * FROM " & sTable & " WHERE " & sChamps & " = '" & txtNoProjSoum.Text & "' AND Type = 'M' AND TypeModif = 'RETOUR' ORDER BY Date, Heure", g_connData, adOpenDynamic, adLockOptimistic)

335     If Not rstProjSoum.EOF Then
          'Ajout de la section "Retour de marchandise"
340       Set itmModif = lvwHistorique.ListItems.Add

345       itmModif.Text = "Retour de marchandise"

350       itmModif.Bold = True

355       Do While Not rstProjSoum.EOF
            'Ajout des retour de marchandise
360         Set itmModif = lvwHistorique.ListItems.Add

365         Call rstEmploye.Open("SELECT Employe FROM GRB_Employés WHERE noEmploye = " & rstProjSoum.Fields("NoEmployé"), g_connData, adOpenDynamic, adLockOptimistic)
            'Employé
370         itmModif.Text = rstEmploye.Fields("Employe")
        
375         Call rstEmploye.Close
        
            'Date
380         itmModif.SubItems(I_COL_MODIF_DATE) = rstProjSoum.Fields("Date")
        
            'Heure
385         itmModif.SubItems(I_COL_MODIF_HEURE) = rstProjSoum.Fields("Heure")
        
390         Call rstProjSoum.MoveNext
395       Loop
400     End If
    
405     Call rstProjSoum.Close

410     Call rstProjSoum.Open("SELECT * FROM " & sTable & " WHERE " & sChamps & " = '" & txtNoProjSoum.Text & "' AND Type = 'M' AND TypeModif = 'FACTURATION' ORDER BY Date, Heure", g_connData, adOpenDynamic, adLockOptimistic)
             
415     If Not rstProjSoum.EOF Then
          'Ajout de la section "Facturation"
420       Set itmModif = lvwHistorique.ListItems.Add
    
425       itmModif.Text = "Facturation"
    
430       itmModif.Bold = True
    
435       Do While Not rstProjSoum.EOF
            'Ajout des modifications
440         Set itmModif = lvwHistorique.ListItems.Add
      
445         Call rstEmploye.Open("SELECT Employe FROM GRB_Employés WHERE noEmploye = " & rstProjSoum.Fields("NoEmployé"), g_connData, adOpenDynamic, adLockOptimistic)
      
            'Employé
450         itmModif.Text = rstEmploye.Fields("Employe")
      
455         Call rstEmploye.Close
      
            'Date
460         itmModif.SubItems(I_COL_MODIF_DATE) = rstProjSoum.Fields("Date")
      
            'Heure
465         itmModif.SubItems(I_COL_MODIF_HEURE) = rstProjSoum.Fields("Heure")

            'Montant
470         itmModif.SubItems(I_COL_MODIF_MONTANT) = rstProjSoum.Fields("Valeur")
      
475         Call rstProjSoum.MoveNext
480       Loop
485     End If
  
490     Call rstProjSoum.Close
495     Set rstProjSoum = Nothing

500     Set rstEmploye = Nothing

505     Exit Sub

AfficherErreur:

510     woups "frmProjSoumMec", "RemplirListViewModifications", Err, Erl
End Sub

Private Sub Cmdajouter_Click()

5       On Error GoTo AfficherErreur

        'Ajoute une soumission
10      Dim rstProjSoum As ADODB.Recordset
15      Dim sNumero     As String
20      Dim sNoProjet   As String
25      Dim bExiste     As Boolean
30      Dim bProjet     As Boolean
35      Dim bContinuer  As Boolean
40      Dim bNoValide   As Boolean
  
        'Affiche le message de saisie selon le type
45      If m_eType = TYPE_PROJET Then
50        sNumero = InputBox("Veuillez entrer le numéro du projet")
55      Else
60        If MsgBox("Voulez-vous créer une nouvelle soumission?" & vbNewLine & _
                    "Oui - Nouvelle soumission" & vbNewLine & _
                    "Non - Copie d'un projet dans une soumission", vbYesNo) = vbYes Then
65          sNumero = InputBox("Veuillez entrer le numéro de la soumission")
70        Else
75          sNumero = InputBox("Veuillez entrer le numéro de la soumission")

80          sNoProjet = InputBox("À partir de quel projet voulez-vous créer cette soumission?")

85          bProjet = True
90        End If
95      End If
    
100     If bProjet = True Then
105       If sNumero <> vbNullString And sNoProjet <> vbNullString Then
110         bContinuer = True
115       End If
120     Else
125       If sNumero <> vbNullString Then
130         bContinuer = True
135       End If
140     End If

145     If bContinuer = True Then
150       Screen.MousePointer = vbHourglass

155       bNoValide = True

160       If ValiderFormatNumeroProjSoum(sNumero) = False Then
165         bNoValide = False
170       End If

175       If bNoValide = True Then
180         If ValiderFormatMecanique(sNumero) = False Then
185           bNoValide = False
190         End If
195       End If

200       If bNoValide = True Then
205         If m_eType = TYPE_PROJET Then
210           If ValiderFormatJobSansSoum(sNumero) = False Then
215             bNoValide = False
220           End If
225         Else
230           If ValiderFormatSoumission(sNumero) = False Then
235             bNoValide = False
240           End If
245         End If
250       End If

255       If bNoValide = False Then
260         Screen.MousePointer = vbDefault

265         Exit Sub
270       End If

275       sNumero = UCase(sNumero)

280       Set rstProjSoum = New ADODB.Recordset
          
          'Ouvre le recordset selon le type
285       Call rstProjSoum.Open("SELECT * FROM GRB_SoumissionElec WHERE IDSoumission = '" & sNumero & "'", g_connData, adOpenDynamic, adLockOptimistic)
        
290       If rstProjSoum.EOF Then
295         bExiste = False
300       Else
305         bExiste = True

310         Call MsgBox("Le numéro " & sNumero & " existe dans les soumissions électriques!", vbOKOnly, "Erreur")
315       End If

320       Call rstProjSoum.Close

325       If bExiste = False Then
330         Call rstProjSoum.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & sNumero & "'", g_connData, adOpenDynamic, adLockOptimistic)
        
335         If rstProjSoum.EOF Then
340           bExiste = False
345         Else
350           bExiste = True

355           Call MsgBox("Le numéro " & sNumero & " existe dans les projets électriques!", vbOKOnly, "Erreur")
360         End If

365         Call rstProjSoum.Close
370       End If

375       If bExiste = False Then
380         Call rstProjSoum.Open("SELECT * FROM GRB_SoumissionMec WHERE IDSoumission = '" & sNumero & "'", g_connData, adOpenDynamic, adLockOptimistic)

385         If rstProjSoum.EOF Then
390           bExiste = False
395         Else
400           bExiste = True

405           Call MsgBox("Le numéro " & sNumero & " existe dans les soumissions mécaniques!", vbOKOnly, "Erreur")
410         End If

415         Call rstProjSoum.Close
420       End If
          
425       If bExiste = False Then
430         Call rstProjSoum.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = '" & sNumero & "'", g_connData, adOpenDynamic, adLockOptimistic)

435         If rstProjSoum.EOF Then
440           bExiste = False
445         Else
450           bExiste = True

455           Call MsgBox("Le numéro " & sNumero & " existe dans les projets mécaniques!", vbOKOnly, "Erreur")
460         End If

465         Call rstProjSoum.Close
470       End If
          
          'Si le projet ou la soumission n'existe pas
475       If bExiste = False Then
            'Pour pouvoir réafficher l'élément affiché avant l'ajout au cas où la personne
            'annule l'ajout
480         Call rstProjSoum.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & sNumero & "'", g_connData, adOpenDynamic, adLockOptimistic)

485         If rstProjSoum.EOF = False Then
490           If rstProjSoum.Fields("Ouvert") = False Then
495             Call MsgBox("Ce numéro est fermé!", vbOKOnly, "Erreur")

500             Call rstProjSoum.Close
505             Set rstProjSoum = Nothing

510             Screen.MousePointer = vbDefault

515             Exit Sub
520           End If
525         End If

530         Call rstProjSoum.Close
535         Set rstProjSoum = Nothing

540         If bProjet = False Then
545           Call InitialiserTempsTaux(True)

550           Call InitialiserNouveauxTaux
    
555           m_sAncienProjSoum = txtNoProjSoum.Text
    
              'Affiche le nouveau numéro
560           txtNoProjSoum.Text = sNumero
    
565           Call InitialiserVariables(sNumero)
  
              'Débarre les champs
570           Call BarrerChamps(False)
      
              'Vide les champs
575           Call ViderChamps
580         Else
585           If VerifierProjet(sNoProjet) = True Then
                'Vide les champs
590             Call ViderChamps

595             txtNoProjSoum.Text = sNumero

600             Call RemplirSoumissionProjet(sNumero, sNoProjet)
605           Else
610             Call MsgBox("Le projet " & sNoProjet & " n'existe pas!", vbOKOnly, "Erreur")

615             Screen.MousePointer = vbDefault

620             Exit Sub
625           End If
630         End If

            'Vide la valeur par défaut si demande sous-section
635         m_sSousSection = vbNullString

640         m_bModeAjout = True
645         m_bModeAffichage = False
               
650         lvwSoumission.Height = lvwSoumission.Height * 0.49
655         lvwSoumission.Top = lvwPieces.Top + lvwPieces.Height + (lvwSoumission.Height * 0.02)
        
            'Met le form en mode ajout/modif
660         Call AfficherControles(MODE_AJOUT_MODIF)
665       End If
670     End If
  
675     Screen.MousePointer = vbDefault

680     Exit Sub

AfficherErreur:

685     woups "frmProjSoumMec", "cmdAjouter_Click", Err, Erl
End Sub

Private Function VerifierProjet(ByVal sNoProjet As String) As Boolean
  
5       On Error GoTo AfficherErreur

10      Dim rstProjet As ADODB.Recordset
15      Dim bExiste   As Boolean

20      Set rstProjet = New ADODB.Recordset

25      Call rstProjet.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)

30      If Not rstProjet.EOF Then
35        bExiste = True
40      Else
45        bExiste = False
50      End If

55      Call rstProjet.Close
60      Set rstProjet = Nothing

65      VerifierProjet = bExiste

70      Exit Function
  
AfficherErreur:

75      woups "frmProjSoumMec", "VerifierProjet", Err, Erl
End Function

Private Sub RemplirSoumissionProjet(ByVal sNoSoumission As String, ByVal sNoProjet As String)

5       On Error GoTo AfficherErreur

        'Affiche le projet ou la soumission choisie
10      Dim rstProjSoum  As ADODB.Recordset
15      Dim rstConfig    As ADODB.Recordset
20      Dim bVariables   As Boolean
25      Dim bTauxHoraire As Boolean
30      Dim bPrixPieces  As Boolean

35      Set rstProjSoum = New ADODB.Recordset
40      Set rstConfig = New ADODB.Recordset

45      If MsgBox("Voulez-vous mettre à jour les variables systèmes?" & vbNewLine & _
                  "-  % Profit" & vbNewLine & _
                  "-  % Commission" & vbNewLine & _
                  "-  % Imprévu", vbYesNo) = vbYes Then
50        bVariables = True
55      Else
60        bVariables = False
65      End If

70      If MsgBox("Voulez-vous mettre à jour les taux horaires?", vbYesNo) = vbYes Then
75        bTauxHoraire = True
80      Else
85        bTauxHoraire = False
90      End If

95      If MsgBox("Voulez-vous mettre à jour le prix des pièces?", vbYesNo) = vbYes Then
100       bPrixPieces = True
105     Else
110       bPrixPieces = False
115     End If

120     Call rstProjSoum.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)

125     If rstProjSoum.Fields("TempsProjBarré") = False Then
130       If Not IsNull(rstProjSoum.Fields("TempsDessinProj")) Then
135         m_sTempsDessin = rstProjSoum.Fields("TempsDessinProj")
140       Else
145         m_sTempsDessin = "0"
150       End If

155       If Not IsNull(rstProjSoum.Fields("TempsCoupeProj")) Then
160         m_sTempsCoupe = rstProjSoum.Fields("TempsCoupeProj")
165       Else
170         m_sTempsCoupe = "0"
175       End If

180       If Not IsNull(rstProjSoum.Fields("TempsMachinageProj")) Then
185         m_sTempsMachinage = rstProjSoum.Fields("TempsMachinageProj")
190       Else
195         m_sTempsMachinage = "0"
200       End If

205       If Not IsNull(rstProjSoum.Fields("TempsSoudureProj")) Then
210         m_sTempsSoudure = rstProjSoum.Fields("TempsSoudureProj")
215       Else
220         m_sTempsSoudure = "0"
225       End If

230       If Not IsNull(rstProjSoum.Fields("TempsAssemblageProj")) Then
235         m_sTempsAssemblage = rstProjSoum.Fields("TempsAssemblageProj")
240       Else
245         m_sTempsAssemblage = "0"
250       End If

255       If Not IsNull(rstProjSoum.Fields("TempsPeintureProj")) Then
260         m_sTempsPeinture = rstProjSoum.Fields("TempsPeintureProj")
265       Else
270         m_sTempsPeinture = "0"
275       End If

280       If Not IsNull(rstProjSoum.Fields("TempsTestProj")) Then
285         m_sTempsTest = rstProjSoum.Fields("TempsTestProj")
290       Else
295         m_sTempsTest = "0"
300       End If

305       If Not IsNull(rstProjSoum.Fields("TempsInstallationProj")) Then
310         m_sTempsInstallation = rstProjSoum.Fields("TempsInstallationProj")
315       Else
320         m_sTempsInstallation = "0"
325       End If

330       If Not IsNull(rstProjSoum.Fields("TempsFormationProj")) Then
335         m_sTempsFormation = rstProjSoum.Fields("TempsFormationProj")
340       Else
345         m_sTempsFormation = "0"
350       End If

355       If Not IsNull(rstProjSoum.Fields("TempsGestionProj")) Then
360         m_sTempsGestion = rstProjSoum.Fields("TempsGestionProj")
365       Else
370         m_sTempsGestion = "0"
375       End If

380       If Not IsNull(rstProjSoum.Fields("TempsShippingProj")) Then
385         m_sTempsShipping = rstProjSoum.Fields("TempsShippingProj")
390       Else
395         m_sTempsShipping = "0"
400       End If
405     Else
410       If Not IsNull(rstProjSoum.Fields("TempsDessinConc")) Then
415         m_sTempsDessin = rstProjSoum.Fields("TempsDessinConc")
420       Else
425         m_sTempsDessin = "0"
430       End If

435       If Not IsNull(rstProjSoum.Fields("TempsCoupeConc")) Then
440         m_sTempsCoupe = rstProjSoum.Fields("TempsCoupeConc")
445       Else
450         m_sTempsCoupe = "0"
455       End If

460       If Not IsNull(rstProjSoum.Fields("TempsMachinageConc")) Then
465         m_sTempsMachinage = rstProjSoum.Fields("TempsMachinageConc")
470       Else
475         m_sTempsMachinage = "0"
480       End If

485       If Not IsNull(rstProjSoum.Fields("TempsSoudureConc")) Then
490         m_sTempsSoudure = rstProjSoum.Fields("TempsSoudureConc")
495       Else
500         m_sTempsSoudure = "0"
505       End If

510       If Not IsNull(rstProjSoum.Fields("TempsAssemblageConc")) Then
515         m_sTempsAssemblage = rstProjSoum.Fields("TempsAssemblageConc")
520       Else
525         m_sTempsAssemblage = "0"
530       End If

535       If Not IsNull(rstProjSoum.Fields("TempsPeintureConc")) Then
540         m_sTempsPeinture = rstProjSoum.Fields("TempsPeintureConc")
545       Else
550         m_sTempsPeinture = "0"
555       End If

560       If Not IsNull(rstProjSoum.Fields("TempsTestConc")) Then
565         m_sTempsTest = rstProjSoum.Fields("TempsTestConc")
570       Else
575         m_sTempsTest = "0"
580       End If

585       If Not IsNull(rstProjSoum.Fields("TempsInstallationConc")) Then
590         m_sTempsInstallation = rstProjSoum.Fields("TempsInstallationConc")
595       Else
600         m_sTempsInstallation = "0"
605       End If

610       If Not IsNull(rstProjSoum.Fields("TempsFormationConc")) Then
615         m_sTempsFormation = rstProjSoum.Fields("TempsFormationConc")
620       Else
625         m_sTempsFormation = "0"
630       End If

635       If Not IsNull(rstProjSoum.Fields("TempsGestionConc")) Then
640         m_sTempsGestion = rstProjSoum.Fields("TempsGestionConc")
645       Else
650         m_sTempsGestion = "0"
655       End If

660       If Not IsNull(rstProjSoum.Fields("TempsShippingConc")) Then
665         m_sTempsShipping = rstProjSoum.Fields("TempsShippingConc")
670       Else
675         m_sTempsShipping = "0"
680       End If
685     End If

690     If Not IsNull(rstProjSoum.Fields("NbrePersonne")) Then
695       m_sNbrePersonne = rstProjSoum.Fields("NbrePersonne")
700     Else
705       m_sNbrePersonne = "0"
710     End If

720     If Not IsNull(rstProjSoum.Fields("TempsHebergement")) Then
725       m_sTempsHebergement = rstProjSoum.Fields("TempsHebergement")
730     Else
735       m_sTempsHebergement = "0"
740     End If

745     If Not IsNull(rstProjSoum.Fields("TempsRepas")) Then
750       m_sTempsRepas = rstProjSoum.Fields("TempsRepas")
755     Else
760       m_sTempsRepas = "0"
765     End If

770     If Not IsNull(rstProjSoum.Fields("TempsTransport")) Then
775       m_sTempsTransport = rstProjSoum.Fields("TempsTransport")
780     Else
785       m_sTempsTransport = "0"
790     End If

795     If Not IsNull(rstProjSoum.Fields("TempsUniteMobile")) Then
800       m_sTempsUniteMobile = rstProjSoum.Fields("TempsUniteMobile")
805     Else
810       m_sTempsUniteMobile = "0"
815     End If

820     m_sPrixEmballage = rstProjSoum.Fields("PrixEmballage")

825     Call rstConfig.Open("SELECT * FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

830     If bTauxHoraire = True Then
835       If Not IsNull(rstConfig.Fields("TauxDessinMec")) Then
840         m_sTauxDessin = rstConfig.Fields("TauxDessinMec")
845       Else
850         m_sTauxDessin = "0"
855       End If

860       If Not IsNull(rstConfig.Fields("TauxCoupe")) Then
865         m_sTauxCoupe = rstConfig.Fields("TauxCoupe")
870       Else
875         m_sTauxCoupe = "0"
880       End If

885       If Not IsNull(rstConfig.Fields("TauxMachinage")) Then
890         m_sTauxMachinage = rstConfig.Fields("TauxMachinage")
895       Else
900         m_sTauxMachinage = "0"
905       End If

910       If Not IsNull(rstConfig.Fields("TauxSoudure")) Then
915         m_sTauxSoudure = rstConfig.Fields("TauxSoudure")
920       Else
925         m_sTauxSoudure = "0"
930       End If

935       If Not IsNull(rstConfig.Fields("TauxAssemblageMec")) Then
940         m_sTauxAssemblage = rstConfig.Fields("TauxAssemblageMec")
945       Else
950         m_sTauxAssemblage = "0"
955       End If

960       If Not IsNull(rstConfig.Fields("TauxPeinture")) Then
965         m_sTauxPeinture = rstConfig.Fields("TauxPeinture")
970       Else
975         m_sTauxPeinture = "0"
980       End If

985       If Not IsNull(rstConfig.Fields("TauxTestMec")) Then
990         m_sTauxTest = rstConfig.Fields("TauxTestMec")
995       Else
1000        m_sTauxTest = "0"
1005      End If

1010      If Not IsNull(rstConfig.Fields("TauxInstallationMec")) Then
1015        m_sTauxInstallation = rstConfig.Fields("TauxInstallationMec")
1020      Else
1025        m_sTauxInstallation = "0"
1030      End If

1035      If Not IsNull(rstConfig.Fields("TauxFormationMec")) Then
1040        m_sTauxFormation = rstConfig.Fields("TauxFormationMec")
1045      Else
1050        m_sTauxFormation = "0"
1055      End If

1060      If Not IsNull(rstConfig.Fields("TauxGestionProjetsMec")) Then
1065        m_sTauxGestion = rstConfig.Fields("TauxGestionProjetsMec")
1070      Else
1075        m_sTauxGestion = "0"
1080      End If

1085      If Not IsNull(rstConfig.Fields("TauxShippingMec")) Then
1090        m_sTauxShipping = rstConfig.Fields("TauxShippingMec")
1095      Else
1100        m_sTauxShipping = "0"
1105      End If

1110      m_sTauxHebergement1 = rstConfig.Fields("Hebergement1")
1115      m_sTauxHebergement2 = rstConfig.Fields("Hebergement2")
1120      m_sTauxRepas = rstConfig.Fields("Repas")
1125      m_sTauxTransport = rstConfig.Fields("Standard")
1130      m_sTauxUniteMobile = rstConfig.Fields("UniteMobile")
1135    Else
1140      If Not IsNull(rstProjSoum.Fields("TauxDessin")) Then
1145        m_sTauxDessin = rstProjSoum.Fields("TauxDessin")
1150      Else
1155        m_sTauxDessin = "0"
1160      End If

1165      If Not IsNull(rstProjSoum.Fields("TauxCoupe")) Then
1170        m_sTauxCoupe = rstProjSoum.Fields("TauxCoupe")
1175      Else
1180        m_sTauxCoupe = "0"
1185      End If

1190      If Not IsNull(rstProjSoum.Fields("TauxMachinage")) Then
1195        m_sTauxMachinage = rstProjSoum.Fields("TauxMachinage")
1200      Else
1205        m_sTauxMachinage = "0"
1210      End If

1215      If Not IsNull(rstProjSoum.Fields("TauxSoudure")) Then
1220        m_sTauxSoudure = rstProjSoum.Fields("TauxSoudure")
1225      Else
1230        m_sTauxSoudure = "0"
1235      End If

1240      If Not IsNull(rstProjSoum.Fields("TauxAssemblage")) Then
1245        m_sTauxAssemblage = rstProjSoum.Fields("TauxAssemblage")
1250      Else
1255        m_sTauxAssemblage = "0"
1260      End If

1265      If Not IsNull(rstProjSoum.Fields("TauxPeinture")) Then
1270        m_sTauxPeinture = rstProjSoum.Fields("TauxPeinture")
1275      Else
1280        m_sTauxPeinture = "0"
1285      End If

1290      If Not IsNull(rstProjSoum.Fields("TauxTest")) Then
1295        m_sTauxTest = rstProjSoum.Fields("TauxTest")
1300      Else
1305        m_sTauxTest = "0"
1310      End If

1315      If Not IsNull(rstProjSoum.Fields("TauxInstallation")) Then
1320        m_sTauxInstallation = rstProjSoum.Fields("TauxInstallation")
1325      Else
1330        m_sTauxInstallation = "0"
1335      End If

1340      If Not IsNull(rstProjSoum.Fields("TauxFormation")) Then
1345        m_sTauxFormation = rstProjSoum.Fields("TauxFormation")
1350      Else
1355        m_sTauxFormation = "0"
1360      End If

1365      If Not IsNull(rstProjSoum.Fields("TauxGestion")) Then
1370        m_sTauxGestion = rstProjSoum.Fields("TauxGestion")
1375      Else
1380        m_sTauxGestion = "0"
1385      End If

1390      If Not IsNull(rstProjSoum.Fields("TauxShipping")) Then
1395        m_sTauxShipping = rstProjSoum.Fields("TauxShipping")
1400      Else
1405        m_sTauxShipping = "0"
1410      End If

1415      If Not IsNull(rstProjSoum.Fields("TauxHebergement1")) Then
1420        m_sTauxHebergement1 = rstProjSoum.Fields("TauxHebergement1")
1425      Else
1430        m_sTauxHebergement1 = "0"
1435      End If

1440      If Not IsNull(rstProjSoum.Fields("TauxHebergement2")) Then
1445        m_sTauxHebergement2 = rstProjSoum.Fields("TauxHebergement2")
1450      Else
1455        m_sTauxHebergement2 = "0"
1460      End If

1465      If Not IsNull(rstProjSoum.Fields("TauxRepas")) Then
1470        m_sTauxRepas = rstProjSoum.Fields("TauxRepas")
1475      Else
1480        m_sTauxRepas = "0"
1485      End If

1490      If Not IsNull(rstProjSoum.Fields("TauxTransport")) Then
1495        m_sTauxTransport = rstProjSoum.Fields("TauxTransport")
1500      Else
1505        m_sTauxTransport = "0"
1510      End If

1515      If Not IsNull(rstProjSoum.Fields("TauxUniteMobile")) Then
1520        m_sTauxUniteMobile = rstProjSoum.Fields("TauxUniteMobile")
1525      Else
1530        m_sTauxUniteMobile = "0"
1535      End If
1540    End If

1545    If bVariables = True Then
1550      m_sProfit = rstConfig.Fields("ProfitMec")
1555      m_sCommission = rstConfig.Fields("Commission")
1560      m_sImprevue = rstConfig.Fields("Imprévus")
1565    Else
1570      m_sProfit = rstProjSoum.Fields("Profit")
1575      m_sCommission = rstProjSoum.Fields("Commission")
1580      m_sImprevue = rstProjSoum.Fields("Imprevue")
1585    End If

1590    Call rstConfig.Close
1595    Set rstConfig = Nothing
                           
1600    txtDescription.Text = rstProjSoum.Fields("Description")
1605    txtNbreManuel.Text = rstProjSoum.Fields("manuel")
1610    txtPrixManuel.Text = rstProjSoum.Fields("total_manuel")

1615    txtTotalPieces.Text = rstProjSoum.Fields("Total_Piece")
1620    txtTotalTemps.Text = rstProjSoum.Fields("Total_Temps")
1625    txtPrixTotal.Text = rstProjSoum.Fields("PrixTotal")
1630    txtImprevus.Text = rstProjSoum.Fields("Total_Imprevue")
1635    txtProfit.Text = rstProjSoum.Fields("total_profit")
1640    txtCommission.Text = rstProjSoum.Fields("total_commission")

1645    If Not IsNull(rstProjSoum.Fields("CheminPhotos")) Then
1650      txtCheminPhotos.Text = rstProjSoum.Fields("CheminPhotos")
1655    Else
1660      txtCheminPhotos.Text = vbNullString
1665    End If

1670    If Not IsNull(rstProjSoum.Fields("MontantForfait")) Then
1675      txtForfait.Text = rstProjSoum.Fields("MontantForfait")

1680      If Not IsNull(rstProjSoum.Fields("InitialeForfait")) Then
1685        lblForfaitInitiale.Caption = rstProjSoum.Fields("InitialeForfait")
1690      Else
1695        lblForfaitInitiale.Caption = ""
1700      End If
1705    Else
1710      txtForfait.Text = ""
1715      lblForfaitInitiale.Caption = ""
1720    End If

1725    Call rstProjSoum.Close
1730    Set rstProjSoum = Nothing

        'Affiche les pieces de la soumission
1735    Call RemplirListViewSoumissionProjet(sNoProjet)

1740    If bPrixPieces = True Then
1745      Call UpdatePieces
1750    End If

1755    m_bModeAffichage = False

1760    Call CalculerPrix

1765    Exit Sub

AfficherErreur:

1770    woups "frmProjSoumMec", "RemplirSoumissionProjet", Err, Erl
End Sub

Private Sub RemplirListViewSoumissionProjet(ByVal sNoProjet As String)

5       On Error GoTo AfficherErreur

        'Remplis les pièces de la soumission avec la BD
10      Dim rstProjSoum   As ADODB.Recordset
15      Dim rstSection    As ADODB.Recordset
20      Dim rstFRS        As ADODB.Recordset
25      Dim itmProjSoum   As ListItem
30      Dim bPremierEnr   As Boolean
35      Dim iOrdreSection As Integer
40      Dim sSousSection  As String
45      Dim sSection      As String
50      Dim lColor        As Long
  
55      Call lvwSoumission.ListItems.Clear
  
60      bPremierEnr = True
  
65      Set rstProjSoum = New ADODB.Recordset
  
70      Call rstProjSoum.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoProjet & "' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
    
75      If m_eLangage = ANGLAIS Then
80        sSection = "NomSectionEN"
85      Else
90        sSection = "NomSectionFR"
95      End If

100     Set rstSection = New ADODB.Recordset
105     Set rstFRS = New ADODB.Recordset

110     Do While Not rstProjSoum.EOF
115       Set itmProjSoum = lvwSoumission.ListItems.Add
            
          'Si c'est le premier enregistrement, il faut ajouter la section et la sous-section
120       If bPremierEnr = True Then
125         iOrdreSection = rstProjSoum.Fields("OrdreSection")
130         sSousSection = rstProjSoum.Fields("SousSection")
              
            'Pour avoir le nom de la section
135         Call rstSection.Open("SELECT " & sSection & " FROM GRB_SoumProjSectionMec WHERE IDSection = " & rstProjSoum.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
        
            'Ajout du nom de la section
140         If Not IsNull(rstSection.Fields(sSection)) Then
145           itmProjSoum.SubItems(I_COL_SOUM_PIECE) = rstSection.Fields(sSection)
150         Else
155           itmProjSoum.SubItems(I_COL_SOUM_PIECE) = vbNullString
160         End If
        
165         itmProjSoum.ListSubItems(I_COL_SOUM_PIECE).Bold = True
       
170         Call ValeurParDefaut(itmProjSoum)
                      
175         Call rstSection.Close
          
180         Set itmProjSoum = lvwSoumission.ListItems.Add
        
            'Ajout du nom de la sous-section
185         If sSousSection = S_PAS_SOUS_SECTION Then
190           itmProjSoum.SubItems(I_COL_SOUM_DESCR) = vbNullString
195         Else
200           itmProjSoum.SubItems(I_COL_SOUM_DESCR) = sSousSection
205         End If
        
            'Le tag ne peut pas être remplis si la colonne est vide
210         itmProjSoum.SubItems(I_COL_SOUM_MANUFACT) = " "
      
215         itmProjSoum.ListSubItems(I_COL_SOUM_MANUFACT).Tag = iOrdreSection
               
220         itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Bold = True
        
225         itmProjSoum.Tag = rstProjSoum.Fields("IDSection")
        
230         Call ValeurParDefaut(itmProjSoum)
        
235         Set itmProjSoum = lvwSoumission.ListItems.Add
        
240         bPremierEnr = False
245       Else
            'Si c'est pas le premier enregistrement, il faut vérifier avec l'ancienne section
250         If iOrdreSection <> rstProjSoum.Fields("OrdreSection") Then
255           iOrdreSection = rstProjSoum.Fields("OrdreSection")
        
260           Call rstSection.Open("SELECT " & sSection & " FROM GRB_SoumProjSectionMec WHERE IDSection = " & rstProjSoum.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
          
265           If Not IsNull(rstSection.Fields(sSection)) Then
270             itmProjSoum.SubItems(I_COL_SOUM_PIECE) = rstSection.Fields(sSection)
275           Else
280             itmProjSoum.SubItems(I_COL_SOUM_PIECE) = vbNullString
285           End If
          
290           itmProjSoum.ListSubItems(I_COL_SOUM_PIECE).Bold = True
          
295           Call ValeurParDefaut(itmProjSoum)
          
300           Call rstSection.Close
                
305           Set itmProjSoum = lvwSoumission.ListItems.Add
          
310           sSousSection = rstProjSoum.Fields("SousSection")
          
315           If sSousSection = S_PAS_SOUS_SECTION Then
320             itmProjSoum.SubItems(I_COL_SOUM_DESCR) = vbNullString
325           Else
330             itmProjSoum.SubItems(I_COL_SOUM_DESCR) = rstProjSoum.Fields("SousSection")
335           End If
          
              'Le tag ne peut pas être remplis si la colonne est vide
340           itmProjSoum.SubItems(I_COL_SOUM_MANUFACT) = " "
        
345           itmProjSoum.ListSubItems(I_COL_SOUM_MANUFACT).Tag = iOrdreSection
          
350           itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Bold = True
          
355           Call ValeurParDefaut(itmProjSoum)
          
360           itmProjSoum.Tag = rstProjSoum.Fields("IDSection")
          
365           Set itmProjSoum = lvwSoumission.ListItems.Add
370         Else
              'il faut vérifier avec l'ancienne sous-section
375           If sSousSection <> rstProjSoum.Fields("SousSection") Then
380             sSousSection = rstProjSoum.Fields("SousSection")
            
385             If sSousSection = S_PAS_SOUS_SECTION Then
390               itmProjSoum.SubItems(I_COL_SOUM_DESCR) = vbNullString
395             Else
400               itmProjSoum.SubItems(I_COL_SOUM_DESCR) = sSousSection
405             End If
                            
410             itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Bold = True
         
415             Call ValeurParDefaut(itmProjSoum)
          
420             itmProjSoum.Tag = rstProjSoum.Fields("IDSection")
            
                'Le tag ne peut pas être remplis si la colonne est vide
425             itmProjSoum.SubItems(I_COL_SOUM_MANUFACT) = " "
        
430             itmProjSoum.ListSubItems(I_COL_SOUM_MANUFACT).Tag = iOrdreSection
          
435             Set itmProjSoum = lvwSoumission.ListItems.Add
440           End If
445         End If
450       End If
  
455       If rstProjSoum.Fields("IDFRS") = 0 And rstProjSoum.Fields("NumItem") <> "Texte" And rstProjSoum.Fields("NumItem") <> "Text" Then
460         lColor = COLOR_MAGENTA
465       Else
470         lColor = COLOR_NOIR
475       End If

          'On met l'ID de la section dans le tag
480       itmProjSoum.Tag = rstProjSoum.Fields("IDSection")
      
485       If rstProjSoum.Fields("Visible") = True Then
490         itmProjSoum.Checked = True
495       Else
500         itmProjSoum.Checked = False
505       End If
      
          'Quantité
510       If Not IsNull(rstProjSoum.Fields("Qté")) Then
515         itmProjSoum.Text = rstProjSoum.Fields("Qté")
520       Else
525         itmProjSoum.Text = vbNullString
530       End If
    
535       If rstProjSoum.Fields("Quoté") = True Then
540         itmProjSoum.Text = itmProjSoum.Text & "*"
545         itmProjSoum.ForeColor = COLOR_VERT
550         itmProjSoum.Bold = True
555       Else
560         itmProjSoum.ForeColor = COLOR_NOIR
565         itmProjSoum.Bold = False
570       End If
      
          'Numéro d'item
575       If Not IsNull(rstProjSoum.Fields("NumItem")) Then
580         If rstProjSoum.Fields("NumItem") = "Texte" Or rstProjSoum.Fields("NumItem") = "Text" Then
585           If m_eLangage = ANGLAIS Then
590             itmProjSoum.SubItems(I_COL_SOUM_PIECE) = "Text"
595           Else
600             itmProjSoum.SubItems(I_COL_SOUM_PIECE) = "Texte"
605           End If
610         Else
615           itmProjSoum.SubItems(I_COL_SOUM_PIECE) = rstProjSoum.Fields("NumItem")
620         End If
625       Else
630         itmProjSoum.SubItems(I_COL_SOUM_PIECE) = vbNullString
635       End If
   
640       itmProjSoum.ListSubItems(I_COL_SOUM_PIECE).ForeColor = lColor
     
          'On met le nom de la sous-section dans le tag du numéro d'item
645       itmProjSoum.ListSubItems(I_COL_SOUM_PIECE).Tag = rstProjSoum.Fields("SousSection")
    
650       If itmProjSoum.SubItems(I_COL_SOUM_PIECE) = "Texte" Or itmProjSoum.SubItems(I_COL_SOUM_PIECE) = "Text" Then
655         itmProjSoum.SubItems(I_COL_SOUM_DESCR) = rstProjSoum.Fields("DESC_FR")
660       Else
665         If m_eLangage = ANGLAIS Then
              'Description en anglais
670           If Not IsNull(rstProjSoum.Fields("DESC_EN")) Then
675             itmProjSoum.SubItems(I_COL_SOUM_DESCR) = rstProjSoum.Fields("DESC_EN")
680           Else
685             itmProjSoum.SubItems(I_COL_SOUM_DESCR) = vbNullString
690           End If

695           'On met la description en francais dans le tag de la description en anglais
700           If Not IsNull(rstProjSoum.Fields("DESC_FR")) Then
705             itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = rstProjSoum.Fields("DESC_FR")
710           Else
715             itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = vbNullString
720           End If
725         Else
              'Description en francais
730           If Not IsNull(rstProjSoum.Fields("DESC_FR")) Then
735             itmProjSoum.SubItems(I_COL_SOUM_DESCR) = rstProjSoum.Fields("DESC_FR")
740           Else
745             itmProjSoum.SubItems(I_COL_SOUM_DESCR) = vbNullString
750           End If

              'On met la description en anglais dans le tag de la description en francais
755           If Not IsNull(rstProjSoum.Fields("Desc_EN")) Then
760             itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = rstProjSoum.Fields("Desc_EN")
765           Else
770             itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = vbNullString
775           End If
780         End If
785       End If
    
790       itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).ForeColor = lColor
      
          'Fabricant
795       If Not IsNull(rstProjSoum.Fields("Manufact")) Then
800         itmProjSoum.SubItems(I_COL_SOUM_MANUFACT) = rstProjSoum.Fields("Manufact")
805       Else
810         itmProjSoum.SubItems(I_COL_SOUM_MANUFACT) = vbNullString
815       End If
    
820       itmProjSoum.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = lColor
      
          'On met l'ordre de la section dans le tag du fabricant
825       itmProjSoum.ListSubItems(I_COL_SOUM_MANUFACT).Tag = iOrdreSection
     
          'Prix listé
830       If Trim(rstProjSoum.Fields("Prix_List")) <> vbNullString Then
835         itmProjSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(rstProjSoum.Fields("Prix_list"), MODE_ARGENT, 4)
840       Else
845         itmProjSoum.SubItems(I_COL_SOUM_PRIX_LIST) = " "
850       End If

855       itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = lColor
 
860       itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = rstProjSoum.Fields("PrixOrigine")
  
          'Escompte
865       If Trim(rstProjSoum.Fields("Escompte")) <> vbNullString Then
870         itmProjSoum.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion(rstProjSoum.Fields("Escompte"), MODE_POURCENT)
875       Else
880         itmProjSoum.SubItems(I_COL_SOUM_ESCOMPTE) = " "
885       End If
 
890       itmProjSoum.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = lColor
          
          'Prix net
895       If Trim(rstProjSoum.Fields("Prix_net")) <> vbNullString Then
900         itmProjSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(rstProjSoum.Fields("Prix_net"), MODE_ARGENT, 4)
905       Else
910         itmProjSoum.SubItems(I_COL_SOUM_PRIX_NET) = " "
915       End If
       
920       itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = lColor
                  
          'Fournisseur
925       If Not IsNull(rstProjSoum.Fields("IDFRS")) And rstProjSoum.Fields("IDFRS") <> "0" Then
930         If itmProjSoum.SubItems(I_COL_SOUM_PIECE) <> "Texte" And itmProjSoum.SubItems(I_COL_SOUM_PIECE) <> "Text" Then
935           Call rstFRS.Open("SELECT NomFournisseur FROM GRB_Fournisseur WHERE IDFRS = " & rstProjSoum.Fields("IDFRS"), g_connData, adOpenDynamic, adLockOptimistic)

              'On affiche le nom dans la colonne
940           itmProjSoum.SubItems(I_COL_SOUM_DISTRIB) = rstFRS.Fields("NomFournisseur")
    
945           itmProjSoum.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = lColor
   
              'On affiche l'Id dans le tag
950           itmProjSoum.ListSubItems(I_COL_SOUM_DISTRIB).Tag = rstProjSoum.Fields("IDFRS")
   
955           Call rstFRS.Close
960         End If
965       Else
970         itmProjSoum.SubItems(I_COL_SOUM_DISTRIB) = vbNullString
975         itmProjSoum.ListSubItems(I_COL_SOUM_DISTRIB).Tag = 0
980       End If
 
          'Prix total
985       If Trim(rstProjSoum.Fields("Prix_total")) <> vbNullString Then
990         itmProjSoum.SubItems(I_COL_SOUM_TOTAL) = Conversion(Round(rstProjSoum.Fields("Prix_total"), 2), MODE_ARGENT)
995       Else
1000        itmProjSoum.SubItems(I_COL_SOUM_TOTAL) = " "
1005      End If

1010      itmProjSoum.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = lColor

          'Profit
1015      If Trim(rstProjSoum.Fields("Profit_argent")) <> vbNullString Then
1020        itmProjSoum.SubItems(I_COL_SOUM_PROFIT) = Conversion(Round(rstProjSoum.Fields("Profit_Argent"), 2), MODE_ARGENT)
1025      Else
1030        itmProjSoum.SubItems(I_COL_SOUM_PROFIT) = " "
1035      End If
 
1040      itmProjSoum.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = lColor

1045      If Not IsNull(rstProjSoum.Fields("Commentaire")) Then
1050        itmProjSoum.SubItems(I_COL_SOUM_COMMENTAIRE) = rstProjSoum.Fields("Commentaire")
1055      Else
1060        itmProjSoum.SubItems(I_COL_SOUM_COMMENTAIRE) = vbNullString
1065      End If

1070      itmProjSoum.ListSubItems(I_COL_SOUM_COMMENTAIRE).ForeColor = lColor
      
1075      Call rstProjSoum.MoveNext
1080    Loop
        
1085    Call rstProjSoum.Close
1090    Set rstProjSoum = Nothing

1095    Set rstFRS = Nothing
1100    Set rstSection = Nothing

1105    Exit Sub

AfficherErreur:

1110    woups "frmProjSoumMec", "RemplirListViewProjSoum", Err, Erl
End Sub

Private Sub RechercherProjSoum(ByVal sNoProjSoum As String)

5       On Error GoTo AfficherErreur

        'Méthode qui recherche une soumission ou un projet dans le combo
        'et qui la sélectionne
10      Dim iCompteur As Integer
  
        'Pour chaque élément du combo
15      For iCompteur = 0 To cmbProjSoum.ListCount - 1
          'Si le texte de l'élément du combo est égal au numéro recherché
20        If cmbProjSoum.LIST(iCompteur) = sNoProjSoum Then
            'On le sélectionne
25          cmbProjSoum.ListIndex = iCompteur
      
30          Exit For
35        End If
40      Next

45      If cmbProjSoum.ListIndex = -1 Then
50        cmbProjSoum.ListIndex = 0
55      End If

60      Exit Sub

AfficherErreur:

65      woups "frmProjSoumMec", "RechercherProjSoum", Err, Erl
End Sub

Private Sub RemplirProjSoum()

5       On Error GoTo AfficherErreur

        'Affiche le projet ou la soumission choisie
10      Dim rstProjSoum As ADODB.Recordset
15      Dim rstSoum     As ADODB.Recordset
20      Dim rstClient   As ADODB.Recordset
25      Dim rstContact  As ADODB.Recordset
    
        'Ouvre le recordset selon le type
30      Set rstProjSoum = New ADODB.Recordset
        
35      If m_eType = TYPE_PROJET Then
40        Call rstProjSoum.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
      
45        If Not IsNull(rstProjSoum.Fields("IDSoumission")) Then
50          txtNoSoumission.Text = rstProjSoum.Fields("IDSoumission")
55        End If

60        If Right$(txtNoProjSoum.Text, 2) >= "60" And Right$(txtNoProjSoum.Text, 2) <= "98" Then
65          If Not IsNull(rstProjSoum.Fields("LiaisonChargeable")) And Trim(rstProjSoum.Fields("LiaisonChargeable")) <> "" Then
70            m_sLiaison = rstProjSoum.Fields("LiaisonChargeable")
75          Else
80            m_sLiaison = ""

85            Do While Trim$(m_sLiaison) = ""
90              m_sLiaison = InputBox("Quelle est l'extention au projet " & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 3) & " auquel ce projet sera lié?")
95            Loop

100           rstProjSoum.Fields("LiaisonChargeable") = m_sLiaison

105           Call rstProjSoum.Update
110         End If
115       End If
120     Else
125       Call rstProjSoum.Open("SELECT * FROM GRB_SoumissionMec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
130     End If
    
        'Recordset pour avoir le nom du client
135     Set rstClient = New ADODB.Recordset
        
140     Call rstClient.Open("SELECT NomClient,IDClient FROM GRB_Client WHERE IDClient = " & rstProjSoum.Fields("IDClient"), g_connData, adOpenDynamic, adLockOptimistic)
    
        'Recordset pour avoir le nom du contact
145     Set rstContact = New ADODB.Recordset
        
150     Call rstContact.Open("SELECT NomContact,IDContact FROM GRB_Contact WHERE IDContact = " & rstProjSoum.Fields("IDContact"), g_connData, adOpenDynamic, adLockOptimistic)
        
155     txtClient.Text = rstClient.Fields("NomClient")
    
160     txtcontact.Text = rstContact.Fields("NomContact")
    
165     Call rstClient.Close
170     Set rstClient = Nothing
    
175     Call rstContact.Close
180     Set rstContact = Nothing
           
185     txtDescription.Text = rstProjSoum.Fields("Description")

190     If Not IsNull(rstProjSoum.Fields("manuel")) Then
195       txtNbreManuel.Text = rstProjSoum.Fields("manuel")
200     Else
205       txtNbreManuel.Text = "0"
210     End If

215     txtPrixManuel.Text = rstProjSoum.Fields("total_manuel")
    
220     txtTotalPieces.Text = rstProjSoum.Fields("Total_Piece")
225     txtTotalTemps.Text = rstProjSoum.Fields("Total_Temps")
230     txtPrixTotal.Text = rstProjSoum.Fields("PrixTotal")
235     txtImprevus.Text = rstProjSoum.Fields("Total_Imprevue")
240     txtProfit.Text = rstProjSoum.Fields("total_profit")
245     txtCommission.Text = rstProjSoum.Fields("total_commission")

250     If Not IsNull(rstProjSoum.Fields("CheminPhotos")) Then
255       txtCheminPhotos.Text = rstProjSoum.Fields("CheminPhotos")
260     Else
265       txtCheminPhotos.Text = vbNullString
270     End If

275     If Not IsNull(rstProjSoum.Fields("MontantForfait")) Then
280       txtForfait.Text = Conversion(rstProjSoum.Fields("MontantForfait"), MODE_ARGENT)

285       If Not IsNull(rstProjSoum.Fields("InitialeForfait")) Then
290         lblForfaitInitiale.Caption = rstProjSoum.Fields("InitialeForfait")
295       Else
300         lblForfaitInitiale.Caption = ""
305       End If
310     Else
315       txtForfait.Text = ""
320       lblForfaitInitiale.Caption = ""
325     End If

330     If m_eType = TYPE_PROJET Then
335       If Not IsNull(rstProjSoum.Fields("PrixRéception")) Then
340         If Trim(rstProjSoum.Fields("PrixRéception")) <> "" Then
345           txtPrixReception.Text = Conversion(rstProjSoum.Fields("PrixRéception"), MODE_ARGENT)
350         Else
355           txtPrixReception.Text = Conversion("0", MODE_ARGENT)
360         End If
365       Else
370         txtPrixReception.Text = Conversion("0", MODE_ARGENT)
375       End If

380       If Not IsNull(rstProjSoum.Fields("IDSoumission")) Then
385         Set rstSoum = New ADODB.Recordset

390         Call rstSoum.Open("SELECT * FROM GRB_SoumissionMec WHERE IDSoumission = '" & rstProjSoum.Fields("IDSoumission") & "'", g_connData, adOpenDynamic, adLockOptimistic)

395         If Not rstSoum.EOF Then
400           If Not IsNull(rstSoum.Fields("PrixTotal")) Then
405             If Trim(rstSoum.Fields("PrixTotal")) <> "" Then
410               txtPrixSoumission.Text = Conversion(rstSoum.Fields("PrixTotal"), MODE_ARGENT)
415             Else
420               txtPrixSoumission.Text = Conversion(0, MODE_ARGENT)
425             End If
430           Else
435             txtPrixSoumission.Text = Conversion(0, MODE_ARGENT)
440           End If
445         Else
450           txtPrixSoumission.Text = Conversion(0, MODE_ARGENT)
455         End If

460         Call rstSoum.Close
465         Set rstSoum = Nothing
470       End If
475     End If
   
480     Call rstProjSoum.Close
485     Set rstProjSoum = Nothing
   
        'Affiche les pieces de la soumission
490     Call RemplirListViewProjSoum(txtNoProjSoum.Text)

495     Exit Sub

AfficherErreur:

500     woups "frmProjSoumMec", "RemplirProjSoum", Err, Erl
End Sub

Private Sub RemplirComboCategoriesPieces()

5       On Error GoTo AfficherErreur

        'Remplir le combo categorie
10      Dim rstCatalogueMec As ADODB.Recordset
  
        'Il faut vider le combo avant de le remplir
15      Call cmbPieces.Clear
      
        'Cette méthode crée un recordset contenant les categorie
        'le nom de toutes les tables de la BD
20      Set rstCatalogueMec = New ADODB.Recordset
        
25      Call rstCatalogueMec.Open("SELECT DISTINCT CATEGORIE FROM GRB_CatalogueMec ORDER BY CATEGORIE", g_connData, adOpenDynamic, adLockOptimistic)
    
        'Tant que ce n'est pas la fin des enregistrements
30      Do While Not rstCatalogueMec.EOF
35        Call cmbPieces.AddItem(rstCatalogueMec.Fields("CATEGORIE"))
      
40        Call rstCatalogueMec.MoveNext
45      Loop
    
50      Call rstCatalogueMec.Close
55      Set rstCatalogueMec = Nothing
  
        'Si le combo n'est pas vide, on sélectionne le premier
60      If cmbPieces.ListCount > 0 Then
65        cmbPieces.ListIndex = 0
70      End If

75      Exit Sub

AfficherErreur:

80      woups "frmProjSoumMec", "RemplirComboCategoriesPieces", Err, Erl
End Sub

Private Sub RemplirComboClients(ByVal sRecherche As String)

5       On Error GoTo AfficherErreur

        'Remplit le combo des clients
10      Dim rstClient As ADODB.Recordset
  
15      Call cmbclient.Clear
  
20      Set rstClient = New ADODB.Recordset
  
25      Call rstClient.Open("SELECT NomClient, IDClient FROM GRB_Client WHERE Instr(1, NomClient, '" & Replace(sRecherche, "'", "''") & "') > 0 AND Supprimé = False", g_connData, adOpenDynamic, adLockOptimistic)
    
30      Do While Not rstClient.EOF
          'on met le nom du client dans le combo
35        Call cmbclient.AddItem(rstClient.Fields("NomClient"))
      
          'on met l'id du client dans l'itemdata
40        cmbclient.ItemData(cmbclient.newIndex) = rstClient.Fields("IDClient")
      
45        Call rstClient.MoveNext
50      Loop
    
55      Call rstClient.Close
60      Set rstClient = Nothing

65      Exit Sub

AfficherErreur:

70      woups "frmProjSoumMec", "RemplirComboClients", Err, Erl
End Sub

Private Sub RemplirComboContacts()

5       On Error GoTo AfficherErreur

        'Remplis le combo des contacts selon le client choisi
10      Dim rstContact As ADODB.Recordset
  
15      Call cmbContact.Clear
  
20      If cmbclient.ListIndex > -1 Then
25        Set rstContact = New ADODB.Recordset

30        Call rstContact.Open("SELECT NomContact, IDContact FROM GRB_Contact INNER JOIN GRB_ContactClient ON GRB_Contact.IDContact = GRB_ContactClient.NoContact WHERE GRB_ContactClient.noClient = " & cmbclient.ItemData(cmbclient.ListIndex), g_connData, adOpenDynamic, adLockOptimistic)
            
          'Si il n'y a aucun contact pour le client choisi
35        If rstContact.EOF Then
            'On ajoute tous les contacts
40          Call rstContact.Close
        
45          Call rstContact.Open("SELECT NomContact, IDContact FROM GRB_Contact WHERE Supprimé = False", g_connData, adOpenDynamic, adLockOptimistic)
50        End If
      
55        Do While Not rstContact.EOF
            'On ajoute le nom du contact dans le combo
60          Call cmbContact.AddItem(rstContact.Fields("NomContact"))
        
            'On ajoute l'id du contact dans l'itemdata du combo
65          cmbContact.ItemData(cmbContact.newIndex) = rstContact.Fields("IDContact")
      
70          Call rstContact.MoveNext
75        Loop
       
80        Call rstContact.Close
85        Set rstContact = Nothing
    
          'Si le combo n'est pas vide, on sélectionne le premier élément
90        If cmbContact.ListCount > 0 Then
95          cmbContact.ListIndex = 0
100       End If
105     End If

110     Exit Sub

AfficherErreur:

115     woups "frmProjSoumMec", "RemplirComboContacts", Err, Erl
End Sub

Private Sub RemplirComboSections()

5       On Error GoTo AfficherErreur

        'Remplis le combo des sections
10      Dim rstSection As ADODB.Recordset
15      Dim sChamps    As String
  
20      Call cmbSections.Clear
  
        'Il faut le remplir selon l'ordre
25      Set rstSection = New ADODB.Recordset
        
30      Call rstSection.Open("SELECT * FROM GRB_SoumProjSectionMec ORDER BY Ordre", g_connData, adOpenDynamic, adLockOptimistic)
    
35      Do While Not rstSection.EOF
          'On met le nom de la section dans le combo
40        If m_eLangage = ANGLAIS Then
45          sChamps = "NomSectionEN"
50        Else
55          sChamps = "NomSectionFR"
60        End If
        
65        If Not IsNull(rstSection.Fields(sChamps)) Then
70          Call cmbSections.AddItem(rstSection.Fields(sChamps))
75        Else
80          Call cmbSections.AddItem(vbNullString)
85        End If
          
          'On met l'id de la section dans l'itemdata du combo
90        cmbSections.ItemData(cmbSections.newIndex) = rstSection.Fields("IDSection")
      
95        Call rstSection.MoveNext
100     Loop
    
105     Call rstSection.Close
110     Set rstSection = Nothing
  
        'Si le combo n'est pas vide, on sélectionne le premier élément
115     If cmbSections.ListCount > 0 Then
120       cmbSections.ListIndex = 0
125     End If

130     Exit Sub

AfficherErreur:

135     woups "frmProjSoumMec", "RemplirComboSections", Err, Erl
End Sub

Private Sub cmdImprimer_Click()

5       On Error GoTo AfficherErreur

10      Dim rstProjSoum As ADODB.Recordset
15      Dim sUser       As String

20      If txtNoProjSoum.Text <> vbNullString Then
25        If VerifierSiOuvert(sUser) = False Then
            'ouvre les tables
30          Set rstProjSoum = New ADODB.Recordset

35          If m_eType = TYPE_PROJET Then
40            If MsgBox("Voulez-vous faire imprimer le projet et tous les extras associés à ce projet?", vbYesNo) = vbYes Then
45              Call rstProjSoum.Open("SELECT * FROM GRB_ProjetMec WHERE Left$(IDProjet, 6) = '" & Left$(txtNoProjSoum.Text, 6) & "' ORDER BY IDProjet", g_connData, adOpenDynamic, adLockOptimistic)
50            Else
55              Call rstProjSoum.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = '" & txtNoProjSoum.Text & "' ORDER BY IDProjet", g_connData, adOpenDynamic, adLockOptimistic)
60            End If
65          Else
70            Call rstProjSoum.Open("SELECT * FROM GRB_SoumissionMec WHERE IDSoumission = '" & txtNoProjSoum.Text & "' ORDER BY IDSoumission", g_connData, adOpenDynamic, adLockOptimistic)
75          End If

76          bTrigger = False
77          intdummie = 0

            '***********************************************************************************
            'AJOUT PAR GAÉTAN GINGRAS 07 FÉVRIER 2010
            '***********************************************************************************
            If MsgBox("Désirez-vous afficher les dates de réception et de commande?", vbYesNo, "Date de réception et de commande") = vbYes Then
                bFlag = True
            Else
                bFlag = False
            End If
            '***********************************************************************************

80          Do While Not rstProjSoum.EOF
85            If m_eType = TYPE_PROJET Then
90              Call CalculerTotalRecordset(rstProjSoum.Fields("IDProjet"))
95            End If

100           Call ImprimerProjSoum(rstProjSoum)
101           If Not intdummie = vbYes Then
105             Call ImprimerListePieces(rstProjSoum)
106           End If
110           Call rstProjSoum.MoveNext
115         Loop
  
120         Call rstProjSoum.Close
125         Set rstProjSoum = Nothing
130       Else
135         If m_eType = TYPE_PROJET Then
140           Call MsgBox("Ce projet est en modification par " & sUser & "!", vbOKOnly, "Erreur")
145         Else
150           Call MsgBox("Cette soumission est en modification par " & sUser & "!", vbOKOnly, "Erreur")
155         End If
160       End If
165     End If

170     Exit Sub

AfficherErreur:

175     woups "frmProjSoumMec", "cmdImprimer_Click", Err, Erl
End Sub

Private Sub ImprimerProjSoum(ByVal rstProjSoum As ADODB.Recordset)

5       On Error GoTo AfficherErreur

        'Impression de la feuille de soumission
10      Dim rstPiece             As ADODB.Recordset
15      Dim rstPrixSoum          As ADODB.Recordset
20      Dim rstTemp              As ADODB.Recordset
25      Dim rstImpProjSoum       As ADODB.Recordset
30      Dim rstSoum              As ADODB.Recordset
35      Dim sOrdreSection        As String
40      Dim iCompteurSoum        As Integer
45      Dim sSousSection         As String
50      Dim sSousSectionRS       As String
55      Dim dblTempsDessin       As Double
60      Dim dblTempsCoupe        As Double
65      Dim dblTempsMachinage    As Double
70      Dim dblTempsSoudure      As Double
75      Dim dblTempsAssemblage   As Double
80      Dim dblTempsPeinture     As Double
85      Dim dblTempsTest         As Double
90      Dim dblTempsInstallation As Double
95      Dim dblTempsFormation    As Double
100     Dim dblTempsGestion      As Double
105     Dim dblTempsShipping     As Double
110     Dim dblTotalTemps        As Double
115     Dim dblTotalAutre        As Double
120     Dim dblTotalReste        As Double
125     Dim dblTotalHebergement  As Double
130     Dim dblTotalRepas        As Double
135     Dim dblTotalTransport    As Double
140     Dim dblTotalUniteMobile  As Double
145     Dim sChampsSection       As String
150     Dim sNoProjet            As String
155     Dim sNoSoumission        As String
160     Dim dblPrixEmballage     As Double
      
        'Supprime les données de l'impression
165     Call g_connData.Execute("DELETE * FROM GRB_impression_soumission")
      
170     iCompteurSoum = 1
  
175     Screen.MousePointer = vbHourglass

180     Set rstImpProjSoum = New ADODB.Recordset

185     Call rstImpProjSoum.Open("SELECT * FROM GRB_impression_soumission", g_connData, adOpenDynamic, adLockOptimistic)
    
190     sOrdreSection = vbNullString
            
195     Set rstPiece = New ADODB.Recordset
            
200     If m_eType = TYPE_PROJET Then
205       sNoProjet = rstProjSoum.Fields("IDProjet")

210       If Not IsNull(rstProjSoum.Fields("IDSoumission")) Then
215         sNoSoumission = rstProjSoum.Fields("IDSoumission")
220       Else
225         sNoSoumission = vbNullString
230       End If

235       Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoProjet & "' And Type = 'M' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
240     Else
245       sNoProjet = vbNullString
250       sNoSoumission = rstProjSoum.Fields("IDSoumission")

255       Call rstPiece.Open("SELECT * FROM GRB_Soumission_Pieces WHERE IDSoumission = '" & sNoSoumission & "' And Type = 'M' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
260     End If
  
265     Set rstTemp = New ADODB.Recordset

270     Do While Not rstPiece.EOF
275       sSousSectionRS = rstPiece.Fields("SousSection")
       
280       If sSousSectionRS = S_PAS_SOUS_SECTION Then
285         sSousSectionRS = " "
290       End If
      
295       If sOrdreSection <> rstPiece.Fields("OrdreSection") Then
            'remplis la table impression_soumission
            'ajoute seulement la section
300         sOrdreSection = rstPiece.Fields("OrdreSection")
        
305         If m_eLangage = ANGLAIS Then
310           sChampsSection = "NomSectionEN"
315         Else
320           sChampsSection = "NomSectionFR"
325         End If

330         Call rstTemp.Open("SELECT " & sChampsSection & " FROM GRB_SoumProjSectionMec WHERE IDSection = " & rstPiece.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
         
            'Ajoute la section dans la soumission
335         Call rstImpProjSoum.AddNew
          
340         rstImpProjSoum.Fields("NoLigne") = iCompteurSoum

345         If m_eType = TYPE_PROJET Then
350           rstImpProjSoum.Fields("IDSoumission") = sNoProjet
355         Else
360           rstImpProjSoum.Fields("IDSoumission") = sNoSoumission
365         End If
        
370         If Not IsNull(rstTemp.Fields(sChampsSection)) Then
375           rstImpProjSoum.Fields("NomSection") = rstTemp.Fields(sChampsSection)
380         Else
385           rstImpProjSoum.Fields("NomSection") = " "
390         End If
           
395         Call rstImpProjSoum.Update
          
400         iCompteurSoum = iCompteurSoum + 1
        
405         Call rstTemp.Close
          
410         sSousSection = rstPiece.Fields("SousSection")
          
415         If sSousSection = S_PAS_SOUS_SECTION Or sSousSection = "" Then
420           sSousSection = " "
425         End If
          
430         Call rstImpProjSoum.AddNew

435         rstImpProjSoum.Fields("NoLigne") = iCompteurSoum

440         If m_eType = TYPE_PROJET Then
445           rstImpProjSoum.Fields("IDSoumission") = sNoProjet
450         Else
455           rstImpProjSoum.Fields("IDSoumission") = sNoSoumission
460         End If

465         rstImpProjSoum.Fields("SousSection") = sSousSection
        
470         Call rstImpProjSoum.Update
          
475         iCompteurSoum = iCompteurSoum + 1
480       Else
            'ajoute une soussection dans impression_soum
485         If sSousSection <> sSousSectionRS Then
490           sSousSection = sSousSectionRS
      
495           Call rstImpProjSoum.AddNew
        
500           rstImpProjSoum.Fields("NoLigne") = iCompteurSoum

505           If m_eType = TYPE_PROJET Then
510             rstImpProjSoum.Fields("IDSoumission") = sNoProjet
515           Else
520             rstImpProjSoum.Fields("IDSoumission") = sNoSoumission
525           End If

530           rstImpProjSoum.Fields("SousSection") = sSousSectionRS
                    
535           Call rstImpProjSoum.Update
          
540           iCompteurSoum = iCompteurSoum + 1
545         End If
550       End If
          
          'ajoute une piece dans impression_soum
555       Call rstImpProjSoum.AddNew
      
560       rstImpProjSoum.Fields("NoLigne") = iCompteurSoum

565       If m_eType = TYPE_PROJET Then
570         rstImpProjSoum.Fields("IDsoumission") = sNoProjet
575       Else
580         rstImpProjSoum.Fields("IDSoumission") = sNoSoumission
585       End If

590       rstImpProjSoum.Fields("NumItem") = rstPiece.Fields("NumItem")
595       rstImpProjSoum.Fields("Qté") = rstPiece.Fields("Qté")
      
600       If m_eLangage = ANGLAIS Then
605         rstImpProjSoum.Fields("Description") = rstPiece.Fields("DESC_EN")
610       Else
615         rstImpProjSoum.Fields("Description") = rstPiece.Fields("DESC_FR")
620       End If
      
        '************************************************************************************************
        'SECTION MODIFIER PAR GAÉTAN GINGRAS LE 07 FÉVRIER 2010
        '************************************************************************************************
        
625       rstImpProjSoum.Fields("MANUFACT") = rstPiece.Fields("MANUFACT")
630       'rstImpProjSoum.Fields("PRIX_LIST") = rstPiece.Fields("PRIX_LIST")
      
635       'If Trim(rstPiece.Fields("ESCOMPTE")) <> vbNullString Then
640       '  rstImpProjSoum.Fields("ESCOMPTE") = Conversion(Replace(rstPiece.Fields("ESCOMPTE"), ".", ","), MODE_POURCENT)
645       'Else
650       '  rstImpProjSoum.Fields("ESCOMPTE") = Conversion(0, MODE_POURCENT)
655       'End If
      
660       rstImpProjSoum.Fields("PRIX_NET") = rstPiece.Fields("PRIX_NET")
                     
665       Call rstTemp.Open("SELECT NomFournisseur FROM GRB_Fournisseur WHERE IDFRS = " & rstPiece.Fields("IDFRS"), g_connData, adOpenDynamic, adLockOptimistic)
           
670       If Not rstTemp.EOF Then
675         rstImpProjSoum.Fields("NomFournisseur") = rstTemp.Fields("NomFournisseur")
680       End If
          
685       Call rstTemp.Close
               
690       rstImpProjSoum.Fields("PRIX_TOTAL") = rstPiece.Fields("PRIX_TOTAL")
695       rstImpProjSoum.Fields("PROFIT_ARGENT") = rstPiece.Fields("PROFIT_ARGENT")


    ''''''''''''''''''''''''''''''''''''''''''''''''''''
        '' ajout du Numéros sequentiel de commande dans impression
    ''''''''''''''''''''''''''''''''
        rstImpProjSoum.Fields("NoSéquentiel") = rstPiece.Fields("NoSéquentiel")
        
        
        
        
        'AJOUT DE CETTE SECTION PAR GAÉTAN GINGRAS LE 6 FÉVRIER 2010
        '*************************************************************
        If m_eType = TYPE_PROJET Then
            If Trim(rstPiece.Fields("DateRéception")) <> vbNullString Then
                rstImpProjSoum.Fields("DateReception") = rstPiece.Fields("DateRéception")
            Else
                rstImpProjSoum.Fields("DateReception") = ""
            End If
            If Trim(rstPiece.Fields("DateCommande")) <> vbNullString Then
                rstImpProjSoum.Fields("DateCommande") = rstPiece.Fields("DateCommande")
            Else
                rstImpProjSoum.Fields("DateCommande") = ""
            End If
        Else    'il n'y a pas de champ date de réception et de commande dans la table GRB_Soumission_Pièces
            rstImpProjSoum.Fields("DateReception") = ""
            rstImpProjSoum.Fields("DateCommande") = ""
        End If
        '************************************************************************************************
        'FIN DE LA SECTION DE MODIFICATION
        '************************************************************************************************


       
700       Call rstImpProjSoum.Update
     
705       iCompteurSoum = iCompteurSoum + 1
    
          'prochaine enreg
710       Call rstPiece.MoveNext
715     Loop
        
        'ferme les tables
717     Call rstImpProjSoum.Close
  
        ''''''''''''''''''''''''''''''''''
        ' rapport soumission, met dans l'ordre de ligne
        ''''''''''''''''''''''''''''''''''''
718     Dim sProjet As String

719     If m_eType = TYPE_PROJET Then
720         sProjet = sNoProjet
721     Else
722         sProjet = sNoSoumission
723     End If

724     Call rstImpProjSoum.Open("SELECT * FROM GRB_impression_Soumission WHERE IDSoumission = '" & sProjet & "' ORDER BY NoLigne", g_connData, adOpenDynamic, adLockOptimistic)

725     Set DR_SoumissionMec.DataSource = rstImpProjSoum

        '***********************************************************************
        'Pour cas d'urgence, la fonction d'exporter dans Excel va être ici.
        'Nous pourrons plus tard le mettre à un meilleur endroit.
        'Gaétan Gingras le 14 mai 2009
        '***********************************************************************

726     If bTrigger = False Then
727         bTrigger = True
728         intdummie = MsgBox("Désirez-vous exporter les données dans Excel, SEULEMENT ?", vbYesNo + vbInformation, "Exportation dans Excel")
729     End If
730     If intdummie = vbYes Then
742         Dim sqlstr As String
743         Dim rstExport As ADODB.Recordset
744         Set rstExport = New ADODB.Recordset
745         sqlstr = "SELECT GRB_impression_soumission.IDSoumission, CDbl([Qté]) AS Quantité, GRB_impression_soumission.NumItem, GRB_impression_soumission.Description, GRB_impression_soumission.Manufact, CDbl([Prix_list]) AS PrixdeListe, CDbl(Left([escompte],Len([escompte])-1)) AS Escomptes, CDbl([Prix_net]) AS prix_nette,GRB_impression_soumission.Prix_total - GRB_impression_soumission.Profit_Argent AS Prix_Total, GRB_impression_soumission.NomFournisseur , GRB_impression_soumission.DateReception, GRB_impression_soumission.DateCommande , GRB_impression_soumission.NoSéquentiel "
746         sqlstr = sqlstr + "FROM GRB_impression_soumission "
747         sqlstr = sqlstr + "WHERE (((GRB_impression_soumission.IDSoumission)='" & sProjet & "') AND ((GRB_impression_soumission.NumItem) Is Not Null)) "
748         sqlstr = sqlstr + "ORDER BY GRB_impression_soumission.noligne"
749         Call rstExport.Open(sqlstr, g_connData, adOpenDynamic, adLockOptimistic)
750         Call ExportdansExcel(rstExport)
751         Screen.MousePointer = vbDefault
752         Exit Sub
753     End If
        '***********************************************************************
        
755     Call TraduireImpressionSoumission

760     If m_eType = TYPE_PROJET Then
765       DR_SoumissionMec.Sections("Section5").Controls("shpCadrePrixReception").Visible = True
770       DR_SoumissionMec.Sections("Section5").Controls("lblTitrePrixReception").Visible = True
775       DR_SoumissionMec.Sections("Section5").Controls("lblPrixReception").Visible = True

780       DR_SoumissionMec.Sections("Section5").Controls("shpCadrePrixSoumission").Visible = True
785       DR_SoumissionMec.Sections("Section5").Controls("lblTitrePrixSoumission").Visible = True
790       DR_SoumissionMec.Sections("Section5").Controls("lblPrixSoumission").Visible = True
795     Else
800       DR_SoumissionMec.Sections("Section5").Controls("shpCadrePrixReception").Visible = False
805       DR_SoumissionMec.Sections("Section5").Controls("lblTitrePrixReception").Visible = False
810       DR_SoumissionMec.Sections("Section5").Controls("lblPrixReception").Visible = False

815       DR_SoumissionMec.Sections("Section5").Controls("shpCadrePrixSoumission").Visible = False
820       DR_SoumissionMec.Sections("Section5").Controls("lblTitrePrixSoumission").Visible = False
825       DR_SoumissionMec.Sections("Section5").Controls("lblPrixSoumission").Visible = False
830     End If
                              
        'affiche la date
        '**************************************************
        'ajout par Gaétan Gingras le 20 mai 2009
834     If MsgBox("Désirez-vous afficher la date en bas de page ?", vbYesNo + vbInformation, "Affichage de la date") = vbYes Then
835         DR_SoumissionMec.Sections("section3").Controls("lbldate").Caption = ConvertDate(Date)
836     Else
837         DR_SoumissionMec.Sections("section3").Controls("lbldate").Caption = " "
838     End If
        '**************************************************
        
        'affiche entete
840     DR_SoumissionMec.Sections("Section2").Controls("lblSoumission").Caption = sNoSoumission
       
845     If m_eType = TYPE_PROJET Then
850       DR_SoumissionMec.Sections("Section2").Controls("lblprojet").Caption = rstProjSoum.Fields("IDProjet")
855     Else
860       DR_SoumissionMec.Sections("Section2").Controls("lblProjet").Caption = vbNullString
865     End If
                
870     DR_SoumissionMec.Sections("Section2").Controls("lbldescription").Caption = rstProjSoum.Fields("Description")
    
875     Call rstTemp.Open("SELECT NomClient FROM GRB_Client WHERE IDClient = " & rstProjSoum.Fields("IDClient"), g_connData, adOpenDynamic, adLockOptimistic)
      
880     DR_SoumissionMec.Sections("Section2").Controls("lblclient").Caption = rstTemp.Fields("NomClient")

885     Call rstTemp.Close
      
890     Call rstTemp.Open("SELECT NomContact FROM GRB_Contact WHERE IDContact = " & rstProjSoum.Fields("IDContact"), g_connData, adOpenDynamic, adLockOptimistic)
      
895     DR_SoumissionMec.Sections("Section2").Controls("lblcontact").Caption = rstTemp.Fields("NomContact")
                   
900     Call rstTemp.Close
      
        'Affiche pied d'état
     
        'Temps
905     If m_eType = TYPE_SOUMISSION Then
910       If Not IsNull(rstProjSoum.Fields("TauxDessin")) Then
915         DR_SoumissionMec.Sections("Section5").Controls("lblTauxDessin").Caption = rstProjSoum.Fields("TauxDessin")
920       Else
925         DR_SoumissionMec.Sections("Section5").Controls("lblTauxDessin").Caption = "0"
930       End If

935       If Not IsNull(rstProjSoum.Fields("TauxCoupe")) Then
940         DR_SoumissionMec.Sections("Section5").Controls("lblTauxCoupe").Caption = rstProjSoum.Fields("TauxCoupe")
945       Else
950         DR_SoumissionMec.Sections("Section5").Controls("lblTauxCoupe").Caption = "0"
955       End If

960       If Not IsNull(rstProjSoum.Fields("TauxMachinage")) Then
965         DR_SoumissionMec.Sections("Section5").Controls("lblTauxMachinage").Caption = rstProjSoum.Fields("TauxMachinage")
970       Else
975         DR_SoumissionMec.Sections("Section5").Controls("lblTauxMachinage").Caption = "0"
980       End If

985       If Not IsNull(rstProjSoum.Fields("TauxSoudure")) Then
990         DR_SoumissionMec.Sections("Section5").Controls("lblTauxSoudure").Caption = rstProjSoum.Fields("TauxSoudure")
995       Else
1000        DR_SoumissionMec.Sections("Section5").Controls("lblTauxSoudure").Caption = "0"
1005      End If

1010      If Not IsNull(rstProjSoum.Fields("TauxAssemblage")) Then
1015        DR_SoumissionMec.Sections("Section5").Controls("lblTauxAssemblage").Caption = rstProjSoum.Fields("TauxAssemblage")
1020      Else
1025        DR_SoumissionMec.Sections("Section5").Controls("lblTauxAssemblage").Caption = "0"
1030      End If

1035      If Not IsNull(rstProjSoum.Fields("TauxPeinture")) Then
1040        DR_SoumissionMec.Sections("Section5").Controls("lblTauxPeinture").Caption = rstProjSoum.Fields("TauxPeinture")
1045      Else
1050        DR_SoumissionMec.Sections("Section5").Controls("lblTauxPeinture").Caption = "0"
1055      End If

1060      If Not IsNull(rstProjSoum.Fields("TauxTest")) Then
1065        DR_SoumissionMec.Sections("Section5").Controls("lblTauxTest").Caption = rstProjSoum.Fields("TauxTest")
1070      Else
1075        DR_SoumissionMec.Sections("Section5").Controls("lblTauxTest").Caption = "0"
1080      End If

1085      If Not IsNull(rstProjSoum.Fields("TauxInstallation")) Then
1090        DR_SoumissionMec.Sections("Section5").Controls("lblTauxInstallation").Caption = rstProjSoum.Fields("TauxInstallation")
1095      Else
1100        DR_SoumissionMec.Sections("Section5").Controls("lblTauxInstallation").Caption = "0"
1105      End If

1110      If Not IsNull(rstProjSoum.Fields("TauxFormation")) Then
1115        DR_SoumissionMec.Sections("Section5").Controls("lblTauxFormation").Caption = rstProjSoum.Fields("TauxFormation")
1120      Else
1125        DR_SoumissionMec.Sections("Section5").Controls("lblTauxFormation").Caption = "0"
1130      End If

1135      If Not IsNull(rstProjSoum.Fields("TauxGestion")) Then
1140        DR_SoumissionMec.Sections("Section5").Controls("lblTauxGestion").Caption = rstProjSoum.Fields("TauxGestion")
1145      Else
1150        DR_SoumissionMec.Sections("Section5").Controls("lblTauxGestion").Caption = "0"
1155      End If

1160      If Not IsNull(rstProjSoum.Fields("TauxShipping")) Then
1165        DR_SoumissionMec.Sections("Section5").Controls("lblTauxShipping").Caption = rstProjSoum.Fields("TauxShipping")
1170      Else
1175        DR_SoumissionMec.Sections("Section5").Controls("lblTauxShipping").Caption = "0"
1180      End If

1185      If Not IsNull(rstProjSoum.Fields("TempsDessin")) Then
1190        DR_SoumissionMec.Sections("Section5").Controls("lblTempsDessinSoum").Caption = rstProjSoum.Fields("TempsDessin")
1195      Else
1200        DR_SoumissionMec.Sections("Section5").Controls("lblTempsDessinSoum").Caption = "0"
1205      End If

1210      If Not IsNull(rstProjSoum.Fields("TempsCoupe")) Then
1215        DR_SoumissionMec.Sections("Section5").Controls("lblTempsCoupeSoum").Caption = rstProjSoum.Fields("TempsCoupe")
1220      Else
1225        DR_SoumissionMec.Sections("Section5").Controls("lblTempsCoupeSoum").Caption = "0"
1230      End If

1235      If Not IsNull(rstProjSoum.Fields("TempsMachinage")) Then
1240        DR_SoumissionMec.Sections("Section5").Controls("lblTempsMachinageSoum").Caption = rstProjSoum.Fields("TempsMachinage")
1245      Else
1250        DR_SoumissionMec.Sections("Section5").Controls("lblTempsMachinageSoum").Caption = "0"
1255      End If

1260      If Not IsNull(rstProjSoum.Fields("TempsSoudure")) Then
1265        DR_SoumissionMec.Sections("Section5").Controls("lblTempsSoudureSoum").Caption = rstProjSoum.Fields("TempsSoudure")
1270      Else
1275        DR_SoumissionMec.Sections("Section5").Controls("lblTempsSoudureSoum").Caption = "0"
1280      End If

1285      If Not IsNull(rstProjSoum.Fields("TempsAssemblage")) Then
1290        DR_SoumissionMec.Sections("Section5").Controls("lblTempsAssemblageSoum").Caption = rstProjSoum.Fields("TempsAssemblage")
1295      Else
1300        DR_SoumissionMec.Sections("Section5").Controls("lblTempsAssemblageSoum").Caption = "0"
1305      End If

1310      If Not IsNull(rstProjSoum.Fields("TempsPeinture")) Then
1315        DR_SoumissionMec.Sections("Section5").Controls("lblTempsPeintureSoum").Caption = rstProjSoum.Fields("TempsPeinture")
1320      Else
1325        DR_SoumissionMec.Sections("Section5").Controls("lblTempsPeintureSoum").Caption = "0"
1330      End If

1335      If Not IsNull(rstProjSoum.Fields("TempsTest")) Then
1340        DR_SoumissionMec.Sections("Section5").Controls("lblTempsTestSoum").Caption = rstProjSoum.Fields("TempsTest")
1345      Else
1350        DR_SoumissionMec.Sections("Section5").Controls("lblTempsTestSoum").Caption = "0"
1355      End If

1360      If Not IsNull(rstProjSoum.Fields("TempsInstallation")) Then
1365        DR_SoumissionMec.Sections("Section5").Controls("lblTempsInstallationSoum").Caption = rstProjSoum.Fields("TempsInstallation")
1370      Else
1375        DR_SoumissionMec.Sections("Section5").Controls("lblTempsInstallationSoum").Caption = "0"
1380      End If

1385      If Not IsNull(rstProjSoum.Fields("TempsFormation")) Then
1390        DR_SoumissionMec.Sections("Section5").Controls("lblTempsFormationSoum").Caption = rstProjSoum.Fields("TempsFormation")
1395      Else
1400        DR_SoumissionMec.Sections("Section5").Controls("lblTempsFormationSoum").Caption = "0"
1405      End If

1410      If Not IsNull(rstProjSoum.Fields("TempsGestion")) Then
1415        DR_SoumissionMec.Sections("Section5").Controls("lblTempsGestionSoum").Caption = rstProjSoum.Fields("TempsGestion")
1420      Else
1425        DR_SoumissionMec.Sections("Section5").Controls("lblTempsGestionSoum").Caption = "0"
1430      End If

1435      If Not IsNull(rstProjSoum.Fields("TempsShipping")) Then
1440        DR_SoumissionMec.Sections("Section5").Controls("lblTempsShippingSoum").Caption = rstProjSoum.Fields("TempsShipping")
1445      Else
1450        DR_SoumissionMec.Sections("Section5").Controls("lblTempsShippingSoum").Caption = "0"
1455      End If

1460      If Not IsNull(rstProjSoum.Fields("TempsDessin")) Then
1465        If IsNumeric(rstProjSoum.Fields("TempsDessin")) Then
1470          dblTempsDessin = CDbl(rstProjSoum.Fields("TempsDessin"))
1475        Else
1480          dblTempsDessin = 0
1485        End If
1490      Else
1495        dblTempsDessin = 0
1500      End If

1505      If Not IsNull(rstProjSoum.Fields("TempsCoupe")) Then
1510        If IsNumeric(rstProjSoum.Fields("TempsCoupe")) Then
1515          dblTempsCoupe = CDbl(rstProjSoum.Fields("TempsCoupe"))
1520        Else
1525          dblTempsCoupe = 0
1530        End If
1535      Else
1540        dblTempsCoupe = 0
1545      End If

1550      If Not IsNull(rstProjSoum.Fields("TempsMachinage")) Then
1555        If IsNumeric(rstProjSoum.Fields("TempsMachinage")) Then
1560          dblTempsMachinage = CDbl(rstProjSoum.Fields("TempsMachinage"))
1565        Else
1570          dblTempsMachinage = 0
1575        End If
1580      Else
1585        dblTempsMachinage = 0
1590      End If

1595      If Not IsNull(rstProjSoum.Fields("TempsSoudure")) Then
1600        If IsNumeric(rstProjSoum.Fields("TempsSoudure")) Then
1605          dblTempsSoudure = CDbl(rstProjSoum.Fields("TempsSoudure"))
1610        Else
1615          dblTempsSoudure = 0
1620        End If
1625      Else
1630        dblTempsSoudure = 0
1635      End If

1640      If Not IsNull(rstProjSoum.Fields("TempsAssemblage")) Then
1645        If IsNumeric(rstProjSoum.Fields("TempsAssemblage")) Then
1650          dblTempsAssemblage = CDbl(rstProjSoum.Fields("TempsAssemblage"))
1655        Else
1660          dblTempsAssemblage = 0
1665        End If
1670      Else
1675        dblTempsAssemblage = 0
1680      End If

1685      If Not IsNull(rstProjSoum.Fields("TempsPeinture")) Then
1690        If IsNumeric(rstProjSoum.Fields("TempsPeinture")) Then
1695          dblTempsPeinture = CDbl(rstProjSoum.Fields("TempsPeinture"))
1700        Else
1705          dblTempsPeinture = 0
1710        End If
1715      Else
1720        dblTempsPeinture = 0
1725      End If

1730      If Not IsNull(rstProjSoum.Fields("TempsTest")) Then
1735        If IsNumeric(rstProjSoum.Fields("TempsTest")) Then
1740          dblTempsTest = CDbl(rstProjSoum.Fields("TempsTest"))
1745        Else
1750          dblTempsTest = 0
1755        End If
1760      Else
1765        dblTempsTest = 0
1770      End If

1775      If Not IsNull(rstProjSoum.Fields("TempsInstallation")) Then
1780        If IsNumeric(rstProjSoum.Fields("TempsInstallation")) Then
1785          dblTempsInstallation = CDbl(rstProjSoum.Fields("TempsInstallation"))
1790        Else
1795          dblTempsInstallation = 0
1800        End If
1805      Else
1810        dblTempsInstallation = 0
1815      End If

1820      If Not IsNull(rstProjSoum.Fields("TempsFormation")) Then
1825        If IsNumeric(rstProjSoum.Fields("TempsFormation")) Then
1830          dblTempsFormation = CDbl(rstProjSoum.Fields("TempsFormation"))
1835        Else
1840          dblTempsFormation = 0
1845        End If
1850      Else
1855        dblTempsFormation = 0
1860      End If

1865      If Not IsNull(rstProjSoum.Fields("TempsGestion")) Then
1870        If IsNumeric(rstProjSoum.Fields("TempsGestion")) Then
1875          dblTempsGestion = CDbl(rstProjSoum.Fields("TempsGestion"))
1880        Else
1885          dblTempsGestion = 0
1890        End If
1895      Else
1900        dblTempsGestion = 0
1905      End If

1910      If Not IsNull(rstProjSoum.Fields("TempsShipping")) Then
1915        If IsNumeric(rstProjSoum.Fields("TempsShipping")) Then
1920          dblTempsShipping = CDbl(rstProjSoum.Fields("TempsShipping"))
1925        Else
1930          dblTempsShipping = 0
1935        End If
1940      Else
1945        dblTempsShipping = 0
1950      End If

1955      dblTotalTemps = dblTempsDessin + _
                          dblTempsCoupe + _
                          dblTempsMachinage + _
                          dblTempsSoudure + _
                          dblTempsAssemblage + _
                          dblTempsPeinture + _
                          dblTempsTest + _
                          dblTempsInstallation + _
                          dblTempsFormation + _
                          dblTempsGestion + _
                          dblTempsShipping
                          
1960      DR_SoumissionMec.Sections("Section5").Controls("lblTotalTempsRHSoum").Caption = dblTotalTemps

1965      DR_SoumissionMec.Sections("Section5").Controls("lblTempsDessinProj").Caption = "---"
1970      DR_SoumissionMec.Sections("Section5").Controls("lblTempsCoupeProj").Caption = "---"
1975      DR_SoumissionMec.Sections("Section5").Controls("lblTempsMachinageProj").Caption = "---"
1980      DR_SoumissionMec.Sections("Section5").Controls("lblTempsSoudureProj").Caption = "---"
1985      DR_SoumissionMec.Sections("Section5").Controls("lblTempsAssemblageProj").Caption = "---"
1990      DR_SoumissionMec.Sections("Section5").Controls("lblTempsPeintureProj").Caption = "---"
1995      DR_SoumissionMec.Sections("Section5").Controls("lblTempsTestProj").Caption = "---"
2000      DR_SoumissionMec.Sections("Section5").Controls("lblTempsInstallationProj").Caption = "---"
2005      DR_SoumissionMec.Sections("Section5").Controls("lblTempsFormationProj").Caption = "---"
2010      DR_SoumissionMec.Sections("Section5").Controls("lblTempsGestionProj").Caption = "---"
2015      DR_SoumissionMec.Sections("Section5").Controls("lblTempsShippingProj").Caption = "---"

2020      DR_SoumissionMec.Sections("Section5").Controls("lblTotalTempsRHProj").Caption = "---"

2025      DR_SoumissionMec.Sections("Section5").Controls("lblTempsDessinConc").Caption = "---"
2030      DR_SoumissionMec.Sections("Section5").Controls("lblTempsCoupeConc").Caption = "---"
2035      DR_SoumissionMec.Sections("Section5").Controls("lblTempsMachinageConc").Caption = "---"
2040      DR_SoumissionMec.Sections("Section5").Controls("lblTempsSoudureConc").Caption = "---"
2045      DR_SoumissionMec.Sections("Section5").Controls("lblTempsAssemblageConc").Caption = "---"
2050      DR_SoumissionMec.Sections("Section5").Controls("lblTempsPeintureConc").Caption = "---"
2055      DR_SoumissionMec.Sections("Section5").Controls("lblTempsTestConc").Caption = "---"
2060      DR_SoumissionMec.Sections("Section5").Controls("lblTempsInstallationConc").Caption = "---"
2065      DR_SoumissionMec.Sections("Section5").Controls("lblTempsFormationConc").Caption = "---"
2070      DR_SoumissionMec.Sections("Section5").Controls("lblTempsGestionConc").Caption = "---"
2075      DR_SoumissionMec.Sections("Section5").Controls("lblTempsShippingConc").Caption = "---"

2080      DR_SoumissionMec.Sections("Section5").Controls("lblTotalTempsRHConc").Caption = "---"
2085    Else
2090      If Not IsNull(rstProjSoum.Fields("TauxDessin")) Then
2095        DR_SoumissionMec.Sections("Section5").Controls("lblTauxDessin").Caption = rstProjSoum.Fields("TauxDessin")
2100      Else
2105        DR_SoumissionMec.Sections("Section5").Controls("lblTauxDessin").Caption = "0"
2110      End If

2115      If Not IsNull(rstProjSoum.Fields("TauxCoupe")) Then
2120        DR_SoumissionMec.Sections("Section5").Controls("lblTauxCoupe").Caption = rstProjSoum.Fields("TauxCoupe")
2125      Else
2130        DR_SoumissionMec.Sections("Section5").Controls("lblTauxCoupe").Caption = "0"
2135      End If

2140      If Not IsNull(rstProjSoum.Fields("TauxMachinage")) Then
2145        DR_SoumissionMec.Sections("Section5").Controls("lblTauxMachinage").Caption = rstProjSoum.Fields("TauxMachinage")
2150      Else
2155        DR_SoumissionMec.Sections("Section5").Controls("lblTauxMachinage").Caption = "0"
2160      End If

2165      If Not IsNull(rstProjSoum.Fields("TauxSoudure")) Then
2170        DR_SoumissionMec.Sections("Section5").Controls("lblTauxSoudure").Caption = rstProjSoum.Fields("TauxSoudure")
2175      Else
2180        DR_SoumissionMec.Sections("Section5").Controls("lblTauxSoudure").Caption = "0"
2185      End If

2190      If Not IsNull(rstProjSoum.Fields("TauxAssemblage")) Then
2195        DR_SoumissionMec.Sections("Section5").Controls("lblTauxAssemblage").Caption = rstProjSoum.Fields("TauxAssemblage")
2200      Else
2205        DR_SoumissionMec.Sections("Section5").Controls("lblTauxAssemblage").Caption = "0"
2210      End If

2215      If Not IsNull(rstProjSoum.Fields("TauxPeinture")) Then
2220        DR_SoumissionMec.Sections("Section5").Controls("lblTauxPeinture").Caption = rstProjSoum.Fields("TauxPeinture")
2225      Else
2230        DR_SoumissionMec.Sections("Section5").Controls("lblTauxPeinture").Caption = "0"
2235      End If

2240      If Not IsNull(rstProjSoum.Fields("TauxTest")) Then
2245        DR_SoumissionMec.Sections("Section5").Controls("lblTauxTest").Caption = rstProjSoum.Fields("TauxTest")
2250      Else
2255        DR_SoumissionMec.Sections("Section5").Controls("lblTauxTest").Caption = "0"
2260      End If

2265      If Not IsNull(rstProjSoum.Fields("TauxInstallation")) Then
2270        DR_SoumissionMec.Sections("Section5").Controls("lblTauxInstallation").Caption = rstProjSoum.Fields("TauxInstallation")
2275      Else
2280        DR_SoumissionMec.Sections("Section5").Controls("lblTauxInstallation").Caption = "0"
2285      End If

2290      If Not IsNull(rstProjSoum.Fields("TauxFormation")) Then
2295        DR_SoumissionMec.Sections("Section5").Controls("lblTauxFormation").Caption = rstProjSoum.Fields("TauxFormation")
2300      Else
2305        DR_SoumissionMec.Sections("Section5").Controls("lblTauxFormation").Caption = "0"
2310      End If

2315      If Not IsNull(rstProjSoum.Fields("TauxGestion")) Then
2320        DR_SoumissionMec.Sections("Section5").Controls("lblTauxGestion").Caption = rstProjSoum.Fields("TauxGestion")
2325      Else
2330        DR_SoumissionMec.Sections("Section5").Controls("lblTauxGestion").Caption = "0"
2335      End If

2340      If Not IsNull(rstProjSoum.Fields("TauxShipping")) Then
2345        DR_SoumissionMec.Sections("Section5").Controls("lblTauxShipping").Caption = rstProjSoum.Fields("TauxShipping")
2350      Else
2355        DR_SoumissionMec.Sections("Section5").Controls("lblTauxShipping").Caption = "0"
2360      End If

2365      If Not IsNull(rstProjSoum.Fields("IDSoumission")) Then
2370        Set rstSoum = New ADODB.Recordset

2375        Call rstSoum.Open("SELECT * FROM GRB_SoumissionMec WHERE IDSoumission = '" & rstProjSoum.Fields("IDSoumission") & "'", g_connData, adOpenDynamic, adLockOptimistic)

2380        If Not rstSoum.EOF Then
2385          If Not IsNull(rstSoum.Fields("TempsDessin")) Then
2390            DR_SoumissionMec.Sections("Section5").Controls("lblTempsDessinSoum").Caption = rstSoum.Fields("TempsDessin")
2395          Else
2400            DR_SoumissionMec.Sections("Section5").Controls("lblTempsDessinSoum").Caption = "0"
2405          End If

2410          If Not IsNull(rstSoum.Fields("TempsCoupe")) Then
2415            DR_SoumissionMec.Sections("Section5").Controls("lblTempsCoupeSoum").Caption = rstSoum.Fields("TempsCoupe")
2420          Else
2425            DR_SoumissionMec.Sections("Section5").Controls("lblTempsCoupeSoum").Caption = "0"
2430          End If

2435          If Not IsNull(rstSoum.Fields("TempsMachinage")) Then
2440            DR_SoumissionMec.Sections("Section5").Controls("lblTempsMachinageSoum").Caption = rstSoum.Fields("TempsMachinage")
2445          Else
2450            DR_SoumissionMec.Sections("Section5").Controls("lblTempsMachinageSoum").Caption = "0"
2455          End If

2460          If Not IsNull(rstSoum.Fields("TempsSoudure")) Then
2465            DR_SoumissionMec.Sections("Section5").Controls("lblTempsSoudureSoum").Caption = rstSoum.Fields("TempsSoudure")
2470          Else
2475            DR_SoumissionMec.Sections("Section5").Controls("lblTempsSoudureSoum").Caption = "0"
2480          End If

2485          If Not IsNull(rstSoum.Fields("TempsAssemblage")) Then
2490            DR_SoumissionMec.Sections("Section5").Controls("lblTempsAssemblageSoum").Caption = rstSoum.Fields("TempsAssemblage")
2495          Else
2500            DR_SoumissionMec.Sections("Section5").Controls("lblTempsAssemblageSoum").Caption = "0"
2505          End If

2510          If Not IsNull(rstSoum.Fields("TempsPeinture")) Then
2515            DR_SoumissionMec.Sections("Section5").Controls("lblTempsPeintureSoum").Caption = rstSoum.Fields("TempsPeinture")
2520          Else
2525            DR_SoumissionMec.Sections("Section5").Controls("lblTempsPeintureSoum").Caption = "0"
2530          End If

2535          If Not IsNull(rstSoum.Fields("TempsTest")) Then
2540            DR_SoumissionMec.Sections("Section5").Controls("lblTempsTestSoum").Caption = rstSoum.Fields("TempsTest")
2545          Else
2550            DR_SoumissionMec.Sections("Section5").Controls("lblTempsTestSoum").Caption = "0"
2555          End If

2560          If Not IsNull(rstSoum.Fields("TempsInstallation")) Then
2565            DR_SoumissionMec.Sections("Section5").Controls("lblTempsInstallationSoum").Caption = rstSoum.Fields("TempsInstallation")
2570          Else
2575            DR_SoumissionMec.Sections("Section5").Controls("lblTempsInstallationSoum").Caption = "0"
2580          End If

2585          If Not IsNull(rstSoum.Fields("TempsFormation")) Then
2590            DR_SoumissionMec.Sections("Section5").Controls("lblTempsFormationSoum").Caption = rstSoum.Fields("TempsFormation")
2595          Else
2600            DR_SoumissionMec.Sections("Section5").Controls("lblTempsFormationSoum").Caption = "0"
2605          End If

2610          If Not IsNull(rstSoum.Fields("TempsGestion")) Then
2615            DR_SoumissionMec.Sections("Section5").Controls("lblTempsGestionSoum").Caption = rstSoum.Fields("TempsGestion")
2620          Else
2625            DR_SoumissionMec.Sections("Section5").Controls("lblTempsGestionSoum").Caption = "0"
2630          End If

2635          If Not IsNull(rstSoum.Fields("TempsShipping")) Then
2640            DR_SoumissionMec.Sections("Section5").Controls("lblTempsShippingSoum").Caption = rstSoum.Fields("TempsShipping")
2645          Else
2650            DR_SoumissionMec.Sections("Section5").Controls("lblTempsShippingSoum").Caption = "0"
2655          End If

2660          If Not IsNull(rstSoum.Fields("TempsDessin")) Then
2665            If IsNumeric(rstSoum.Fields("TempsDessin")) Then
2670              dblTempsDessin = CDbl(rstSoum.Fields("TempsDessin"))
2675            Else
2680              dblTempsDessin = 0
2685            End If
2690          Else
2695            dblTempsDessin = 0
2700          End If

2705          If Not IsNull(rstSoum.Fields("TempsCoupe")) Then
2710            If IsNumeric(rstSoum.Fields("TempsCoupe")) Then
2715              dblTempsCoupe = CDbl(rstSoum.Fields("TempsCoupe"))
2720            Else
2725              dblTempsCoupe = 0
2730            End If
2735          Else
2740            dblTempsCoupe = 0
2745          End If

2750          If Not IsNull(rstSoum.Fields("TempsMachinage")) Then
2755            If IsNumeric(rstSoum.Fields("TempsMachinage")) Then
2760              dblTempsMachinage = CDbl(rstSoum.Fields("TempsMachinage"))
2765            Else
2770              dblTempsMachinage = 0
2775            End If
2780          Else
2785            dblTempsMachinage = 0
2790          End If

2795          If Not IsNull(rstSoum.Fields("TempsSoudure")) Then
2800            If IsNumeric(rstSoum.Fields("TempsSoudure")) Then
2805              dblTempsSoudure = CDbl(rstSoum.Fields("TempsSoudure"))
2810            Else
2815              dblTempsSoudure = 0
2820            End If
2825          Else
2830            dblTempsSoudure = 0
2835          End If

2840          If Not IsNull(rstSoum.Fields("TempsAssemblage")) Then
2845            If IsNumeric(rstSoum.Fields("TempsAssemblage")) Then
2850              dblTempsAssemblage = CDbl(rstSoum.Fields("TempsAssemblage"))
2855            Else
2860              dblTempsAssemblage = 0
2865            End If
2870          Else
2875            dblTempsAssemblage = 0
2880          End If

2885          If Not IsNull(rstSoum.Fields("TempsPeinture")) Then
2890            If IsNumeric(rstSoum.Fields("TempsPeinture")) Then
2895              dblTempsPeinture = CDbl(rstSoum.Fields("TempsPeinture"))
2900            Else
2905              dblTempsPeinture = 0
2910            End If
2915          Else
2920            dblTempsPeinture = 0
2925          End If

2930          If Not IsNull(rstSoum.Fields("TempsTest")) Then
2935            If IsNumeric(rstSoum.Fields("TempsTest")) Then
2940              dblTempsTest = CDbl(rstSoum.Fields("TempsTest"))
2945            Else
2950              dblTempsTest = 0
2955            End If
2960          Else
2965            dblTempsTest = 0
2970          End If

2975          If Not IsNull(rstSoum.Fields("TempsInstallation")) Then
2980            If IsNumeric(rstSoum.Fields("TempsInstallation")) Then
2985              dblTempsInstallation = CDbl(rstSoum.Fields("TempsInstallation"))
2990            Else
2995              dblTempsInstallation = 0
3000            End If
3005          Else
3010            dblTempsInstallation = 0
3015          End If

3020          If Not IsNull(rstSoum.Fields("TempsFormation")) Then
3025            If IsNumeric(rstSoum.Fields("TempsFormation")) Then
3030              dblTempsFormation = CDbl(rstSoum.Fields("TempsFormation"))
3035            Else
3040              dblTempsFormation = 0
3045            End If
3050          Else
3055            dblTempsFormation = 0
3060          End If

3065          If Not IsNull(rstSoum.Fields("TempsGestion")) Then
3070            If IsNumeric(rstSoum.Fields("TempsGestion")) Then
3075              dblTempsGestion = CDbl(rstSoum.Fields("TempsGestion"))
3080            Else
3085              dblTempsGestion = 0
3090            End If
3095          Else
3100            dblTempsGestion = 0
3105          End If

3110          If Not IsNull(rstSoum.Fields("TempsShipping")) Then
3115            If IsNumeric(rstSoum.Fields("TempsShipping")) Then
3120              dblTempsShipping = CDbl(rstSoum.Fields("TempsShipping"))
3125            Else
3130              dblTempsShipping = 0
3135            End If
3140          Else
3145            dblTempsShipping = 0
3150          End If
  
3155          dblTotalTemps = dblTempsDessin + _
                              dblTempsCoupe + _
                              dblTempsMachinage + _
                              dblTempsSoudure + _
                              dblTempsAssemblage + _
                              dblTempsPeinture + _
                              dblTempsTest + _
                              dblTempsInstallation + _
                              dblTempsFormation + _
                              dblTempsGestion + _
                              dblTempsShipping
                          
3160          DR_SoumissionMec.Sections("Section5").Controls("lblTotalTempsRHSoum").Caption = dblTotalTemps
3165        Else
3170          DR_SoumissionMec.Sections("Section5").Controls("lblTempsDessinSoum").Caption = "---"
3175          DR_SoumissionMec.Sections("Section5").Controls("lblTempsCoupeSoum").Caption = "---"
3180          DR_SoumissionMec.Sections("Section5").Controls("lblTempsMachinageSoum").Caption = "---"
3185          DR_SoumissionMec.Sections("Section5").Controls("lblTempsSoudureSoum").Caption = "---"
3190          DR_SoumissionMec.Sections("Section5").Controls("lblTempsAssemblageSoum").Caption = "---"
3195          DR_SoumissionMec.Sections("Section5").Controls("lblTempsPeintureSoum").Caption = "---"
3200          DR_SoumissionMec.Sections("Section5").Controls("lblTempsTestSoum").Caption = "---"
3205          DR_SoumissionMec.Sections("Section5").Controls("lblTempsInstallationSoum").Caption = "---"
3210          DR_SoumissionMec.Sections("Section5").Controls("lblTempsFormationSoum").Caption = "---"
3215          DR_SoumissionMec.Sections("Section5").Controls("lblTempsGestionSoum").Caption = "---"
3220          DR_SoumissionMec.Sections("Section5").Controls("lblTempsShippingSoum").Caption = "---"

3225          DR_SoumissionMec.Sections("Section5").Controls("lblTotalTempsRHSoum").Caption = "---"
3230        End If

3235        Call rstSoum.Close
3240        Set rstSoum = Nothing
3245      Else
3250        DR_SoumissionMec.Sections("Section5").Controls("lblTempsDessinSoum").Caption = "---"
3255        DR_SoumissionMec.Sections("Section5").Controls("lblTempsCoupeSoum").Caption = "---"
3260        DR_SoumissionMec.Sections("Section5").Controls("lblTempsMachinageSoum").Caption = "---"
3265        DR_SoumissionMec.Sections("Section5").Controls("lblTempsSoudureSoum").Caption = "---"
3270        DR_SoumissionMec.Sections("Section5").Controls("lblTempsAssemblageSoum").Caption = "---"
3275        DR_SoumissionMec.Sections("Section5").Controls("lblTempsPeintureSoum").Caption = "---"
3280        DR_SoumissionMec.Sections("Section5").Controls("lblTempsTestSoum").Caption = "---"
3285        DR_SoumissionMec.Sections("Section5").Controls("lblTempsInstallationSoum").Caption = "---"
3290        DR_SoumissionMec.Sections("Section5").Controls("lblTempsFormationSoum").Caption = "---"
3295        DR_SoumissionMec.Sections("Section5").Controls("lblTempsGestionSoum").Caption = "---"
3300        DR_SoumissionMec.Sections("Section5").Controls("lblTempsShippingSoum").Caption = "---"

3305        DR_SoumissionMec.Sections("Section5").Controls("lblTotalTempsRHSoum").Caption = "---"
3310      End If

3315      If Not IsNull(rstProjSoum.Fields("TempsDessinProj")) Then
3320        DR_SoumissionMec.Sections("Section5").Controls("lblTempsDessinProj").Caption = rstProjSoum.Fields("TempsDessinProj")
3325      Else
3330        DR_SoumissionMec.Sections("Section5").Controls("lblTempsDessinProj").Caption = "0"
3335      End If

3340      If Not IsNull(rstProjSoum.Fields("TempsCoupeProj")) Then
3345        DR_SoumissionMec.Sections("Section5").Controls("lblTempsCoupeProj").Caption = rstProjSoum.Fields("TempsCoupeProj")
3350      Else
3355        DR_SoumissionMec.Sections("Section5").Controls("lblTempsCoupeProj").Caption = "0"
3360      End If

3365      If Not IsNull(rstProjSoum.Fields("TempsMachinageProj")) Then
3370        DR_SoumissionMec.Sections("Section5").Controls("lblTempsMachinageProj").Caption = rstProjSoum.Fields("TempsMachinageProj")
3375      Else
3380        DR_SoumissionMec.Sections("Section5").Controls("lblTempsMachinageProj").Caption = "0"
3385      End If

3390      If Not IsNull(rstProjSoum.Fields("TempsSoudureProj")) Then
3395        DR_SoumissionMec.Sections("Section5").Controls("lblTempsSoudureProj").Caption = rstProjSoum.Fields("TempsSoudureProj")
3400      Else
3405        DR_SoumissionMec.Sections("Section5").Controls("lblTempsSoudureProj").Caption = "0"
3410      End If

3415      If Not IsNull(rstProjSoum.Fields("TempsAssemblageProj")) Then
3420        DR_SoumissionMec.Sections("Section5").Controls("lblTempsAssemblageProj").Caption = rstProjSoum.Fields("TempsAssemblageProj")
3425      Else
3430        DR_SoumissionMec.Sections("Section5").Controls("lblTempsAssemblageProj").Caption = "0"
3435      End If

3440      If Not IsNull(rstProjSoum.Fields("TempsPeintureProj")) Then
3445        DR_SoumissionMec.Sections("Section5").Controls("lblTempsPeintureProj").Caption = rstProjSoum.Fields("TempsPeintureProj")
3450      Else
3455        DR_SoumissionMec.Sections("Section5").Controls("lblTempsPeintureProj").Caption = "0"
3460      End If

3465      If Not IsNull(rstProjSoum.Fields("TempsTestProj")) Then
3470        DR_SoumissionMec.Sections("Section5").Controls("lblTempsTestProj").Caption = rstProjSoum.Fields("TempsTestProj")
3475      Else
3480        DR_SoumissionMec.Sections("Section5").Controls("lblTempsTestProj").Caption = "0"
3485      End If

3490      If Not IsNull(rstProjSoum.Fields("TempsInstallationProj")) Then
3495        DR_SoumissionMec.Sections("Section5").Controls("lblTempsInstallationProj").Caption = rstProjSoum.Fields("TempsInstallationProj")
3500      Else
3505        DR_SoumissionMec.Sections("Section5").Controls("lblTempsInstallationProj").Caption = "0"
3510      End If

3515      If Not IsNull(rstProjSoum.Fields("TempsFormationProj")) Then
3520        DR_SoumissionMec.Sections("Section5").Controls("lblTempsFormationProj").Caption = rstProjSoum.Fields("TempsFormationProj")
3525      Else
3530        DR_SoumissionMec.Sections("Section5").Controls("lblTempsFormationProj").Caption = "0"
3535      End If

3540      If Not IsNull(rstProjSoum.Fields("TempsGestionProj")) Then
3545        DR_SoumissionMec.Sections("Section5").Controls("lblTempsGestionProj").Caption = rstProjSoum.Fields("TempsGestionProj")
3550      Else
3555        DR_SoumissionMec.Sections("Section5").Controls("lblTempsGestionProj").Caption = "0"
3560      End If

3565      If Not IsNull(rstProjSoum.Fields("TempsShippingProj")) Then
3570        DR_SoumissionMec.Sections("Section5").Controls("lblTempsShippingProj").Caption = rstProjSoum.Fields("TempsShippingProj")
3575      Else
3580        DR_SoumissionMec.Sections("Section5").Controls("lblTempsShippingProj").Caption = "0"
3585      End If

3590      If Not IsNull(rstProjSoum.Fields("TempsDessinProj")) Then
3595        If IsNumeric(rstProjSoum.Fields("TempsDessinProj")) Then
3600          dblTempsDessin = CDbl(rstProjSoum.Fields("TempsDessinProj"))
3605        Else
3610          dblTempsDessin = 0
3615        End If
3620      Else
3625        dblTempsDessin = 0
3630      End If

3635      If Not IsNull(rstProjSoum.Fields("TempsCoupeProj")) Then
3640        If IsNumeric(rstProjSoum.Fields("TempsCoupeProj")) Then
3645          dblTempsCoupe = CDbl(rstProjSoum.Fields("TempsCoupeProj"))
3650        Else
3655          dblTempsCoupe = 0
3660        End If
3665      Else
3670        dblTempsCoupe = 0
3675      End If

3680      If Not IsNull(rstProjSoum.Fields("TempsMachinageProj")) Then
3685        If IsNumeric(rstProjSoum.Fields("TempsMachinageProj")) Then
3690          dblTempsMachinage = CDbl(rstProjSoum.Fields("TempsMachinageProj"))
3695        Else
3700          dblTempsMachinage = 0
3705        End If
3710      Else
3715        dblTempsMachinage = 0
3720      End If

3725      If Not IsNull(rstProjSoum.Fields("TempsSoudureProj")) Then
3730        If IsNumeric(rstProjSoum.Fields("TempsSoudureProj")) Then
3735          dblTempsSoudure = CDbl(rstProjSoum.Fields("TempsSoudureProj"))
3740        Else
3745          dblTempsSoudure = 0
3750        End If
3755      Else
3760        dblTempsSoudure = 0
3765      End If

3770      If Not IsNull(rstProjSoum.Fields("TempsAssemblageProj")) Then
3775        If IsNumeric(rstProjSoum.Fields("TempsAssemblageProj")) Then
3780          dblTempsAssemblage = CDbl(rstProjSoum.Fields("TempsAssemblageProj"))
3785        Else
3790          dblTempsAssemblage = 0
3795        End If
3800      Else
3805        dblTempsAssemblage = 0
3810      End If

3815      If Not IsNull(rstProjSoum.Fields("TempsPeintureProj")) Then
3820        If IsNumeric(rstProjSoum.Fields("TempsPeintureProj")) Then
3825          dblTempsPeinture = CDbl(rstProjSoum.Fields("TempsPeintureProj"))
3830        Else
3835          dblTempsPeinture = 0
3840        End If
3845      Else
3850        dblTempsPeinture = 0
3855      End If

3860      If Not IsNull(rstProjSoum.Fields("TempsTestProj")) Then
3865        If IsNumeric(rstProjSoum.Fields("TempsTestProj")) Then
3870          dblTempsTest = CDbl(rstProjSoum.Fields("TempsTestProj"))
3875        Else
3880          dblTempsTest = 0
3885        End If
3890      Else
3895        dblTempsTest = 0
3900      End If

3905      If Not IsNull(rstProjSoum.Fields("TempsInstallationProj")) Then
3910        If IsNumeric(rstProjSoum.Fields("TempsInstallationProj")) Then
3915          dblTempsInstallation = CDbl(rstProjSoum.Fields("TempsInstallationProj"))
3920        Else
3925          dblTempsInstallation = 0
3930        End If
3935      Else
3940        dblTempsInstallation = 0
3945      End If

3950      If Not IsNull(rstProjSoum.Fields("TempsFormationProj")) Then
3955        If IsNumeric(rstProjSoum.Fields("TempsFormationProj")) Then
3960          dblTempsFormation = CDbl(rstProjSoum.Fields("TempsFormationProj"))
3965        Else
3970          dblTempsFormation = 0
3975        End If
3980      Else
3985        dblTempsFormation = 0
3990      End If

3995      If Not IsNull(rstProjSoum.Fields("TempsGestionProj")) Then
4000        If IsNumeric(rstProjSoum.Fields("TempsGestionProj")) Then
4005          dblTempsGestion = CDbl(rstProjSoum.Fields("TempsGestionProj"))
4010        Else
4015          dblTempsGestion = 0
4020        End If
4025      Else
4030        dblTempsGestion = 0
4035      End If

4040      If Not IsNull(rstProjSoum.Fields("TempsShippingProj")) Then
4045        If IsNumeric(rstProjSoum.Fields("TempsShippingProj")) Then
4050          dblTempsShipping = CDbl(rstProjSoum.Fields("TempsShippingProj"))
4055        Else
4060          dblTempsShipping = 0
4065        End If
4070      Else
4075        dblTempsShipping = 0
4080      End If

4085      dblTotalTemps = dblTempsDessin + _
                          dblTempsCoupe + _
                          dblTempsMachinage + _
                          dblTempsSoudure + _
                          dblTempsAssemblage + _
                          dblTempsPeinture + _
                          dblTempsTest + _
                          dblTempsInstallation + _
                          dblTempsFormation + _
                          dblTempsGestion + _
                          dblTempsShipping
                                                    
4090      DR_SoumissionMec.Sections("Section5").Controls("lblTotalTempsRHProj").Caption = dblTotalTemps

4095      If rstProjSoum.Fields("TempsProjBarré") = True Then
4100        If Not IsNull(rstProjSoum.Fields("TempsDessinConc")) Then
4105          DR_SoumissionMec.Sections("Section5").Controls("lblTempsDessinConc").Caption = rstProjSoum.Fields("TempsDessinConc")
4110        Else
4115          DR_SoumissionMec.Sections("Section5").Controls("lblTempsDessinConc").Caption = "0"
4120        End If

4125        If Not IsNull(rstProjSoum.Fields("TempsCoupeConc")) Then
4130          DR_SoumissionMec.Sections("Section5").Controls("lblTempsCoupeConc").Caption = rstProjSoum.Fields("TempsCoupeConc")
4135        Else
4140          DR_SoumissionMec.Sections("Section5").Controls("lblTempsCoupeConc").Caption = "0"
4145        End If

4150        If Not IsNull(rstProjSoum.Fields("TempsMachinageConc")) Then
4155          DR_SoumissionMec.Sections("Section5").Controls("lblTempsMachinageConc").Caption = rstProjSoum.Fields("TempsMachinageConc")
4160        Else
4165          DR_SoumissionMec.Sections("Section5").Controls("lblTempsMachinageConc").Caption = "0"
4170        End If

4175        If Not IsNull(rstProjSoum.Fields("TempsSoudureConc")) Then
4180          DR_SoumissionMec.Sections("Section5").Controls("lblTempsSoudureConc").Caption = rstProjSoum.Fields("TempsSoudureConc")
4185        Else
4190          DR_SoumissionMec.Sections("Section5").Controls("lblTempsSoudureConc").Caption = "0"
4195        End If

4200        If Not IsNull(rstProjSoum.Fields("TempsAssemblageConc")) Then
4205          DR_SoumissionMec.Sections("Section5").Controls("lblTempsAssemblageConc").Caption = rstProjSoum.Fields("TempsAssemblageConc")
4210        Else
4215          DR_SoumissionMec.Sections("Section5").Controls("lblTempsAssemblageConc").Caption = "0"
4220        End If

4225        If Not IsNull(rstProjSoum.Fields("TempsPeintureConc")) Then
4230          DR_SoumissionMec.Sections("Section5").Controls("lblTempsPeintureConc").Caption = rstProjSoum.Fields("TempsPeintureConc")
4235        Else
4240          DR_SoumissionMec.Sections("Section5").Controls("lblTempsPeintureConc").Caption = "0"
4245        End If

4250        If Not IsNull(rstProjSoum.Fields("TempsTestConc")) Then
4255          DR_SoumissionMec.Sections("Section5").Controls("lblTempsTestConc").Caption = rstProjSoum.Fields("TempsTestConc")
4260        Else
4265          DR_SoumissionMec.Sections("Section5").Controls("lblTempsTestConc").Caption = "0"
4270        End If

4275        If Not IsNull(rstProjSoum.Fields("TempsInstallationConc")) Then
4280          DR_SoumissionMec.Sections("Section5").Controls("lblTempsInstallationConc").Caption = rstProjSoum.Fields("TempsInstallationConc")
4285        Else
4290          DR_SoumissionMec.Sections("Section5").Controls("lblTempsInstallationConc").Caption = "0"
4295        End If

4300        If Not IsNull(rstProjSoum.Fields("TempsFormationConc")) Then
4305          DR_SoumissionMec.Sections("Section5").Controls("lblTempsFormationConc").Caption = rstProjSoum.Fields("TempsFormationConc")
4310        Else
4315          DR_SoumissionMec.Sections("Section5").Controls("lblTempsFormationConc").Caption = "0"
4320        End If

4325        If Not IsNull(rstProjSoum.Fields("TempsGestionConc")) Then
4330          DR_SoumissionMec.Sections("Section5").Controls("lblTempsGestionConc").Caption = rstProjSoum.Fields("TempsGestionConc")
4335        Else
4340          DR_SoumissionMec.Sections("Section5").Controls("lblTempsGestionConc").Caption = "0"
4345        End If

4350        If Not IsNull(rstProjSoum.Fields("TempsShippingConc")) Then
4355          DR_SoumissionMec.Sections("Section5").Controls("lblTempsShippingConc").Caption = rstProjSoum.Fields("TempsShippingConc")
4360        Else
4365          DR_SoumissionMec.Sections("Section5").Controls("lblTempsShippingConc").Caption = "0"
4370        End If
  
4375        If Not IsNull(rstProjSoum.Fields("TempsDessinConc")) Then
4380          If IsNumeric(rstProjSoum.Fields("TempsDessinConc")) Then
4385            dblTempsDessin = CDbl(rstProjSoum.Fields("TempsDessinConc"))
4390          Else
4395            dblTempsDessin = 0
4400          End If
4405        Else
4410          dblTempsDessin = 0
4415        End If
  
4420        If Not IsNull(rstProjSoum.Fields("TempsCoupeConc")) Then
4425          If IsNumeric(rstProjSoum.Fields("TempsCoupeConc")) Then
4430            dblTempsCoupe = CDbl(rstProjSoum.Fields("TempsCoupeConc"))
4435          Else
4440            dblTempsCoupe = 0
4445          End If
4450        Else
4455          dblTempsCoupe = 0
4460        End If

4465        If Not IsNull(rstProjSoum.Fields("TempsMachinageConc")) Then
4470          If IsNumeric(rstProjSoum.Fields("TempsMachinageConc")) Then
4475            dblTempsMachinage = CDbl(rstProjSoum.Fields("TempsMachinageConc"))
4480          Else
4485            dblTempsMachinage = 0
4490          End If
4495        Else
4500          dblTempsMachinage = 0
4505        End If

4510        If Not IsNull(rstProjSoum.Fields("TempsSoudureConc")) Then
4515          If IsNumeric(rstProjSoum.Fields("TempsSoudureConc")) Then
4520            dblTempsSoudure = CDbl(rstProjSoum.Fields("TempsSoudureConc"))
4525          Else
4530            dblTempsSoudure = 0
4535          End If
4540        Else
4545          dblTempsSoudure = 0
4550        End If

4555        If Not IsNull(rstProjSoum.Fields("TempsAssemblageConc")) Then
4560          If IsNumeric(rstProjSoum.Fields("TempsAssemblageConc")) Then
4565            dblTempsAssemblage = CDbl(rstProjSoum.Fields("TempsAssemblageConc"))
4570          Else
4575            dblTempsAssemblage = 0
4580          End If
4585        Else
4590          dblTempsAssemblage = 0
4595        End If

4600        If Not IsNull(rstProjSoum.Fields("TempsPeintureConc")) Then
4605          If IsNumeric(rstProjSoum.Fields("TempsPeintureConc")) Then
4610            dblTempsPeinture = CDbl(rstProjSoum.Fields("TempsPeintureConc"))
4615          Else
4620            dblTempsPeinture = 0
4625          End If
4630        Else
4635          dblTempsPeinture = 0
4640        End If

4645        If Not IsNull(rstProjSoum.Fields("TempsTestConc")) Then
4650          If IsNumeric(rstProjSoum.Fields("TempsTestConc")) Then
4655            dblTempsTest = CDbl(rstProjSoum.Fields("TempsTestConc"))
4660          Else
4665            dblTempsTest = 0
4670          End If
4675        Else
4680          dblTempsTest = 0
4685        End If
  
4690        If Not IsNull(rstProjSoum.Fields("TempsInstallationConc")) Then
4695          If IsNumeric(rstProjSoum.Fields("TempsInstallationConc")) Then
4700            dblTempsInstallation = CDbl(rstProjSoum.Fields("TempsInstallationConc"))
4705          Else
4710            dblTempsInstallation = 0
4715          End If
4720        Else
4725          dblTempsInstallation = 0
4730        End If
  
4735        If Not IsNull(rstProjSoum.Fields("TempsFormationConc")) Then
4740          If IsNumeric(rstProjSoum.Fields("TempsFormationConc")) Then
4745            dblTempsFormation = CDbl(rstProjSoum.Fields("TempsFormationConc"))
4750          Else
4755            dblTempsFormation = 0
4760          End If
4765        Else
4770          dblTempsFormation = 0
4775        End If
  
4780        If Not IsNull(rstProjSoum.Fields("TempsGestionConc")) Then
4785          If IsNumeric(rstProjSoum.Fields("TempsGestionConc")) Then
4790            dblTempsGestion = CDbl(rstProjSoum.Fields("TempsGestionConc"))
4795          Else
4800            dblTempsGestion = 0
4805          End If
4810        Else
4815          dblTempsGestion = 0
4820        End If
  
4825        If Not IsNull(rstProjSoum.Fields("TempsShippingConc")) Then
4830          If IsNumeric(rstProjSoum.Fields("TempsShippingConc")) Then
4835            dblTempsShipping = CDbl(rstProjSoum.Fields("TempsShippingConc"))
4840          Else
4845            dblTempsShipping = 0
4850          End If
4855        Else
4860          dblTempsShipping = 0
4865        End If
  
4870        dblTotalTemps = dblTempsDessin + _
                            dblTempsCoupe + _
                            dblTempsMachinage + _
                            dblTempsSoudure + _
                            dblTempsAssemblage + _
                            dblTempsPeinture + _
                            dblTempsTest + _
                            dblTempsInstallation + _
                            dblTempsFormation + _
                            dblTempsGestion + _
                            dblTempsShipping
                            
4875        DR_SoumissionMec.Sections("Section5").Controls("lblTotalTempsRHConc").Caption = dblTotalTemps
4880      Else
4885        DR_SoumissionMec.Sections("Section5").Controls("lblTempsDessinConc").Caption = "---"
4890        DR_SoumissionMec.Sections("Section5").Controls("lblTempsCoupeConc").Caption = "---"
4895        DR_SoumissionMec.Sections("Section5").Controls("lblTempsMachinageConc").Caption = "---"
4900        DR_SoumissionMec.Sections("Section5").Controls("lblTempsSoudureConc").Caption = "---"
4905        DR_SoumissionMec.Sections("Section5").Controls("lblTempsAssemblageConc").Caption = "---"
4910        DR_SoumissionMec.Sections("Section5").Controls("lblTempsPeintureConc").Caption = "---"
4915        DR_SoumissionMec.Sections("Section5").Controls("lblTempsTestConc").Caption = "---"
4920        DR_SoumissionMec.Sections("Section5").Controls("lblTempsInstallationConc").Caption = "---"
4925        DR_SoumissionMec.Sections("Section5").Controls("lblTempsFormationConc").Caption = "---"
4930        DR_SoumissionMec.Sections("Section5").Controls("lblTempsGestionConc").Caption = "---"
4935        DR_SoumissionMec.Sections("Section5").Controls("lblTempsShippingConc").Caption = "---"

4940        DR_SoumissionMec.Sections("Section5").Controls("lblTotalTempsRHConc").Caption = "---"
4945      End If

4950      Call CalculerTempsReelsImpression(rstProjSoum.Fields("IDProjet"))
4955    End If
      
        'Autres frais
4960    If m_eType = TYPE_PROJET Then
4965      DR_SoumissionMec.Sections("Section5").Controls("lblNbrePersonne").Caption = "0"
4970      DR_SoumissionMec.Sections("Section5").Controls("lblTempsHebergement").Caption = "0"
4975      DR_SoumissionMec.Sections("Section5").Controls("lblTauxHebergement1").Caption = "0"
4980      DR_SoumissionMec.Sections("Section5").Controls("lblTauxHebergement2").Caption = "0"
4985      DR_SoumissionMec.Sections("Section5").Controls("lblTempsRepas").Caption = "0"
4990      DR_SoumissionMec.Sections("Section5").Controls("lblTauxRepas").Caption = "0"
4995      DR_SoumissionMec.Sections("Section5").Controls("lblTempsTransport").Caption = "0"
5000      DR_SoumissionMec.Sections("Section5").Controls("lblTauxTransport").Caption = "0"
5005      DR_SoumissionMec.Sections("Section5").Controls("lblTempsUniteMobile").Caption = "0"
5010      DR_SoumissionMec.Sections("Section5").Controls("lblTauxUniteMobile").Caption = "0"
5015    Else
5020      If Not IsNull(rstProjSoum.Fields("NbrePersonne")) Then
5025        DR_SoumissionMec.Sections("Section5").Controls("lblNbrePersonne").Caption = rstProjSoum.Fields("NbrePersonne")
5030      Else
5035        DR_SoumissionMec.Sections("Section5").Controls("lblNbrePersonne").Caption = "0"
5040      End If

5045      If Not IsNull(rstProjSoum.Fields("TempsHebergement")) Then
5050        DR_SoumissionMec.Sections("Section5").Controls("lblTempsHebergement").Caption = rstProjSoum.Fields("TempsHebergement")
5055      Else
5060        DR_SoumissionMec.Sections("Section5").Controls("lblTempsHebergement").Caption = "0"
5065      End If

5070      If Not IsNull(rstProjSoum.Fields("TauxHebergement1")) Then
5075        DR_SoumissionMec.Sections("Section5").Controls("lblTauxHebergement1").Caption = rstProjSoum.Fields("TauxHebergement1")
5080      Else
5085        DR_SoumissionMec.Sections("Section5").Controls("lblTauxHebergement1").Caption = "0"
5090      End If

5095      If Not IsNull(rstProjSoum.Fields("TauxHebergement2")) Then
5100        DR_SoumissionMec.Sections("Section5").Controls("lblTauxHebergement2").Caption = rstProjSoum.Fields("TauxHebergement2")
5105      Else
5110        DR_SoumissionMec.Sections("Section5").Controls("lblTauxHebergement2").Caption = "0"
5115      End If
                  
5120      If Not IsNull(rstProjSoum.Fields("TempsRepas")) Then
5125        DR_SoumissionMec.Sections("Section5").Controls("lblTempsRepas").Caption = rstProjSoum.Fields("TempsRepas")
5130      Else
5135        DR_SoumissionMec.Sections("Section5").Controls("lblTempsRepas").Caption = "0"
5140      End If

5145      If Not IsNull(rstProjSoum.Fields("TauxRepas")) Then
5150        DR_SoumissionMec.Sections("Section5").Controls("lblTauxRepas").Caption = rstProjSoum.Fields("TauxRepas")
5155      Else
5160        DR_SoumissionMec.Sections("Section5").Controls("lblTauxRepas").Caption = "0"
5165      End If

5170      If Not IsNull(rstProjSoum.Fields("TempsTransport")) Then
5175        DR_SoumissionMec.Sections("Section5").Controls("lblTempsTransport").Caption = rstProjSoum.Fields("TempsTransport")
5180      Else
5185        DR_SoumissionMec.Sections("Section5").Controls("lblTempsTransport").Caption = "0"
5190      End If

5195      If Not IsNull(rstProjSoum.Fields("TauxTransport")) Then
5200        DR_SoumissionMec.Sections("Section5").Controls("lblTauxTransport").Caption = rstProjSoum.Fields("TauxTransport")
5205      Else
5210        DR_SoumissionMec.Sections("Section5").Controls("lblTauxTransport").Caption = "0"
5215      End If

5220      If Not IsNull(rstProjSoum.Fields("TempsUniteMobile")) Then
5225        DR_SoumissionMec.Sections("Section5").Controls("lblTempsUniteMobile").Caption = rstProjSoum.Fields("TempsUniteMobile")
5230      Else
5235        DR_SoumissionMec.Sections("Section5").Controls("lblTempsUniteMobile").Caption = "0"
5240      End If

5245      If Not IsNull(rstProjSoum.Fields("TauxUniteMobile")) Then
5250        DR_SoumissionMec.Sections("Section5").Controls("lblTauxUniteMobile").Caption = rstProjSoum.Fields("TauxUniteMobile")
5255      Else
5260        DR_SoumissionMec.Sections("Section5").Controls("lblTauxUniteMobile").Caption = "0"
5265      End If
5270    End If

5275    DR_SoumissionMec.Sections("Section5").Controls("lblPrixEmballage").Caption = rstProjSoum.Fields("PrixEmballage")
5280    DR_SoumissionMec.Sections("Section5").Controls("lblPrixManuel").Caption = rstProjSoum.Fields("Total_Manuel")

5285    DR_SoumissionMec.Sections("Section5").Controls("lblTotalTempsAR").Caption = Conversion(rstProjSoum.Fields("total_temps"), MODE_ARGENT)
5290    DR_SoumissionMec.Sections("Section5").Controls("lblTotalPieceAR").Caption = Conversion(rstProjSoum.Fields("total_piece"), MODE_ARGENT)
5295    DR_SoumissionMec.Sections("Section5").Controls("lblProfit").Caption = Conversion((rstProjSoum.Fields("profit") - 1) * 100, MODE_POURCENT)
5300    DR_SoumissionMec.Sections("Section5").Controls("lblTotalProfit").Caption = Conversion(rstProjSoum.Fields("total_profit"), MODE_ARGENT)
5305    DR_SoumissionMec.Sections("Section5").Controls("lblImprevue").Caption = Conversion(rstProjSoum.Fields("imprevue"), MODE_POURCENT)
5310    DR_SoumissionMec.Sections("Section5").Controls("lblImprevueAR").Caption = Conversion(rstProjSoum.Fields("total_imprevue"), MODE_ARGENT)
5315    DR_SoumissionMec.Sections("Section5").Controls("lblCommission").Caption = Conversion(rstProjSoum.Fields("commission"), MODE_POURCENT)
5320    DR_SoumissionMec.Sections("Section5").Controls("lblCommissionAR").Caption = Conversion(rstProjSoum.Fields("total_commission"), MODE_ARGENT)

5325    If m_eType = TYPE_PROJET Then
5330      If Not IsNull(rstProjSoum.Fields("PrixRéception")) Then
5335        DR_SoumissionMec.Sections("Section5").Controls("lblPrixReception").Caption = Conversion(rstProjSoum.Fields("PrixRéception"), MODE_ARGENT)
5340      Else
5345        DR_SoumissionMec.Sections("Section5").Controls("lblPrixReception").Caption = Conversion(0, MODE_ARGENT)
5350      End If

5355      If Not IsNull(rstProjSoum.Fields("IDSoumission")) Then
5360        Set rstPrixSoum = New ADODB.Recordset

5365        Call rstPrixSoum.Open("SELECT PrixTotal FROM GRB_SoumissionMec WHERE IDSoumission = '" & rstProjSoum.Fields("IDSoumission") & "'", g_connData, adOpenDynamic, adLockOptimistic)

5370        If Not rstPrixSoum.EOF Then
5375          If Not IsNull(rstPrixSoum.Fields("PrixTotal")) Then
5380            DR_SoumissionMec.Sections("Section5").Controls("lblPrixSoumission").Caption = Conversion(rstPrixSoum.Fields("PrixTotal"), MODE_ARGENT)
5385          Else
5390            DR_SoumissionMec.Sections("Section5").Controls("lblPrixSoumission").Caption = Conversion("0", MODE_ARGENT)
5395          End If
5400        Else
5405          DR_SoumissionMec.Sections("Section5").Controls("lblPrixSoumission").Caption = Conversion("0", MODE_ARGENT)
5410        End If

5415        Call rstPrixSoum.Close
5420        Set rstPrixSoum = Nothing
5425      Else
5430        DR_SoumissionMec.Sections("Section5").Controls("lblPrixSoumission").Caption = Conversion("0", MODE_ARGENT)
5435      End If
5440    End If

5445    If m_eType = TYPE_PROJET Then
5450      dblTotalHebergement = 0
5455      dblTotalRepas = 0
5460      dblTotalTransport = 0
5465      dblTotalUniteMobile = 0
5470    Else
5475      If Not IsNull(rstProjSoum.Fields("TotalHebergement")) Then
5480        dblTotalHebergement = CDbl(rstProjSoum.Fields("TotalHebergement"))
5485      Else
5490        dblTotalHebergement = 0
5495      End If
      
5500      If Not IsNull(rstProjSoum.Fields("TotalRepas")) Then
5505        dblTotalRepas = CDbl(rstProjSoum.Fields("TotalRepas"))
5510      Else
5515        dblTotalRepas = 0
5520      End If

5525      If Not IsNull(rstProjSoum.Fields("TempsTransport")) And Not IsNull(rstProjSoum.Fields("TauxTransport")) Then
5530        dblTotalTransport = CDbl(rstProjSoum.Fields("TempsTransport")) * CDbl(rstProjSoum.Fields("TauxTransport"))
5535      Else
5540        dblTotalTransport = 0
5545      End If

5550      If Not IsNull(rstProjSoum.Fields("TempsUniteMobile")) And Not IsNull(rstProjSoum.Fields("TauxUniteMobile")) Then
5555        dblTotalUniteMobile = CDbl(rstProjSoum.Fields("TempsUniteMobile")) * CDbl(rstProjSoum.Fields("TauxUniteMobile"))
5560      Else
5565        dblTotalUniteMobile = 0
5570      End If
5575    End If

5580    dblPrixEmballage = CDbl(Replace(rstProjSoum.Fields("PrixEmballage"), ".", ","))
       
5585    dblTotalReste = dblTotalHebergement + dblTotalRepas + dblTotalTransport + dblTotalUniteMobile + dblPrixEmballage

5590    dblTotalAutre = dblTotalReste + CDbl(rstProjSoum.Fields("total_manuel"))
    
5595    DR_SoumissionMec.Sections("Section5").Controls("lblAutre").Caption = Conversion(CStr(dblTotalAutre), MODE_ARGENT)
    
5600    DR_SoumissionMec.Sections("Section5").Controls("lblGrandTotalAR").Caption = Conversion(rstProjSoum.Fields("prixtotal"), MODE_ARGENT)
            
5605    If rstProjSoum.Fields("MontantForfait") <> "" Then
5610      DR_SoumissionMec.Sections("Section5").Controls("shpCadreForfait").Visible = True
5615      DR_SoumissionMec.Sections("Section5").Controls("lblTitreForfait").Visible = True
5620      DR_SoumissionMec.Sections("Section5").Controls("lblForfait").Visible = True

5625      DR_SoumissionMec.Sections("Section5").Controls("lblTitreForfait").Caption = DR_SoumissionMec.Sections("Section5").Controls("lblTitreForfait").Caption & " ( " & rstProjSoum.Fields("InitialeForfait") & " )"
5630      DR_SoumissionMec.Sections("Section5").Controls("lblForfait").Caption = rstProjSoum.Fields("MontantForfait")
5635    Else
5640      DR_SoumissionMec.Sections("Section5").Controls("shpCadreForfait").Visible = False
5645      DR_SoumissionMec.Sections("Section5").Controls("lblTitreForfait").Visible = False
5650      DR_SoumissionMec.Sections("Section5").Controls("lblForfait").Visible = False
5655    End If

        '************************************************************************************************
        'SECTION MODIFIÉ PAR GAÉTAN GINGRAS LE 07 FÉVRIER 2010
        '************************************************************************************************
        If bFlag = True Then
            DR_SoumissionMec.Sections("Section2").Controls("lbl_DateCommande").Visible = True
            DR_SoumissionMec.Sections("Section2").Controls("lbl_DateReception").Visible = True
            DR_SoumissionMec.Sections("Section1").Controls("txt_DateCommande").Visible = True
            DR_SoumissionMec.Sections("Section1").Controls("txt_DateReception").Visible = True
        Else
            DR_SoumissionMec.Sections("Section2").Controls("lbl_DateCommande").Visible = False
            DR_SoumissionMec.Sections("Section2").Controls("lbl_DateReception").Visible = False
            DR_SoumissionMec.Sections("Section1").Controls("txt_DateCommande").Visible = False
            DR_SoumissionMec.Sections("Section1").Controls("txt_DateReception").Visible = False
        End If
        '************************************************************************************************
        'FIN DE LA SECTION MODIFIÉ
        '************************************************************************************************
           
        'Affiche le rapport soumission
5660    DR_SoumissionMec.Orientation = rptOrientLandscape
    
5665    Call DR_SoumissionMec.Show(vbModal)
             
5670    Call rstImpProjSoum.Close
5675    Set rstImpProjSoum = Nothing

5680    Set rstTemp = Nothing
    
5685    Screen.MousePointer = vbDefault

5690    Exit Sub

AfficherErreur:

5695    woups "frmProjSoumMec", "ImprimerProjSoum", Err, Erl
End Sub

Private Sub ImprimerListePieces(ByVal rstProjSoum As ADODB.Recordset)

5       On Error GoTo AfficherErreur

        'Impression de la feuille de soumission
10      Dim rstPiece            As ADODB.Recordset
15      Dim rstTemp             As ADODB.Recordset
20      Dim rstImpListePiece    As ADODB.Recordset
25      Dim iCompteurPiece      As Integer
30      Dim sSousSection        As String
35      Dim sSection            As String
40      Dim sNoProjet           As String
45      Dim sNoSoumission       As String
50      Dim bAjouterSection     As Boolean
55      Dim bAjouterSousSection As Boolean
60      Dim bAjouterPiece       As Boolean
      
65      Call g_connData.Execute("DELETE * FROM GRB_impression_listepiece")
      
70      iCompteurPiece = 1

75      Screen.MousePointer = vbHourglass
            
80      Set rstPiece = New ADODB.Recordset
            
85      If m_eType = TYPE_PROJET Then
90        sNoProjet = rstProjSoum.Fields("IDProjet")

95        If Not IsNull(rstProjSoum.Fields("IDSoumission")) Then
100         sNoSoumission = rstProjSoum.Fields("IDSoumission")
105       End If

110       Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoProjet & "' And Type = 'M' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
115     Else
120       sNoProjet = vbNullString
125       sNoSoumission = rstProjSoum.Fields("IDSoumission")

130       Call rstPiece.Open("SELECT * FROM GRB_Soumission_Pieces WHERE IDSoumission = '" & sNoSoumission & "' And Type = 'M' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
135     End If
    
140     Set rstImpListePiece = New ADODB.Recordset
145     Set rstTemp = New ADODB.Recordset

150     rstImpListePiece.CursorLocation = adUseClient
    
155     Do While Not rstPiece.EOF
160       If rstPiece.Fields("Visible") = True Then
165         bAjouterSection = True
170         bAjouterSousSection = True
175         bAjouterPiece = True

180         rstImpListePiece.Filter = ""

185         Call rstImpListePiece.Open("SELECT * FROM GRB_impression_listepiece WHERE IDSection = '" & rstPiece.Fields("IDSection") & "'", g_connData, adOpenDynamic, adLockOptimistic)

190         If Not rstImpListePiece.EOF Then
195           bAjouterSection = False

200           Do While Not rstImpListePiece.EOF
205             If rstImpListePiece.Fields("NomSousSection") = rstPiece.Fields("SousSection") Then
210               bAjouterSousSection = False

215               If rstPiece.Fields("NumItem") <> "Texte" And rstPiece.Fields("NumItem") <> "Text" Then
220                 If rstImpListePiece.Fields("NumItem") = rstPiece.Fields("NumItem") Then
225                   bAjouterPiece = False

230                   rstImpListePiece.Fields("Qté") = CDbl(Replace(rstImpListePiece.Fields("Qté"), ".", ",")) + CDbl(rstPiece.Fields("Qté"))

235                   Call rstImpListePiece.Update

240                   If rstImpListePiece.Fields("Qté") = 0 Then
245                     Call rstImpListePiece.Delete

250                     rstImpListePiece.Filter = "NomSousSection = '" & Replace(rstPiece.Fields("SousSection"), "'", "''") & "'"

255                     If rstImpListePiece.RecordCount = 1 Then
260                       Call rstImpListePiece.Delete

265                       rstImpListePiece.Filter = "IDSection = '" & rstPiece.Fields("IDSection") & "'"

270                       If rstImpListePiece.RecordCount = 1 Then
275                         Call rstImpListePiece.Delete
280                       End If
285                     End If
290                   End If
295                 End If
300               Else
305                 Exit Do
310               End If
315             End If

320             Call rstImpListePiece.MoveNext
325           Loop
330         End If

335         If bAjouterSection = True Then
340           If m_eLangage = ANGLAIS Then
345             sSection = "NomSectionEN"
350           Else
355             sSection = "NomSectionFR"
360           End If

365           Call rstTemp.Open("SELECT " & sSection & " FROM GRB_SoumProjSectionMec WHERE IDSection = " & rstPiece.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
          
370           Call rstImpListePiece.AddNew
          
375           rstImpListePiece("NoLigne") = iCompteurPiece
380           rstImpListePiece("IDSoumission") = sNoSoumission
         
385           If Not IsNull(rstTemp.Fields(sSection)) Then
390             rstImpListePiece.Fields("Section") = rstTemp.Fields(sSection)
395           Else
400             rstImpListePiece.Fields("Section") = " "
405           End If

410           rstImpListePiece.Fields("IDSection") = rstPiece.Fields("IDSection")
          
415           Call rstImpListePiece.Update
                   
420           iCompteurPiece = iCompteurPiece + 1
          
425           Call rstTemp.Close
430         End If
          
435         If bAjouterSousSection = True Then
440           sSousSection = rstPiece.Fields("SousSection")
          
445           If sSousSection = S_PAS_SOUS_SECTION Then
450             sSousSection = " "
455           End If
         
460           Call rstImpListePiece.AddNew
          
465           rstImpListePiece.Fields("NoLigne") = iCompteurPiece
470           rstImpListePiece.Fields("IDSoumission") = sNoSoumission
475           rstImpListePiece.Fields("SousSection") = sSousSection
480           rstImpListePiece.Fields("NomSousSection") = rstPiece.Fields("SousSection")
485           rstImpListePiece.Fields("IDSection") = rstPiece.Fields("IDSection")
          
490           Call rstImpListePiece.Update
          
495           iCompteurPiece = iCompteurPiece + 1
500         End If
              
505         If bAjouterPiece = True Then
              'ajoute une piece dans la liste de pièce
510           Call rstImpListePiece.AddNew
      
515           rstImpListePiece.Fields("NoLigne") = iCompteurPiece
520           rstImpListePiece.Fields("IDSoumission") = sNoSoumission
525           rstImpListePiece.Fields("numitem") = rstPiece.Fields("numitem")
530           rstImpListePiece.Fields("qté") = rstPiece.Fields("qté")
        
535           If m_eLangage = ANGLAIS Then
540             rstImpListePiece.Fields("description") = rstPiece.Fields("DESC_EN")
545           Else
550             rstImpListePiece.Fields("description") = rstPiece.Fields("DESC_FR")
555           End If
       
560           rstImpListePiece.Fields("manufact") = rstPiece.Fields("manufact")

565           rstImpListePiece.Fields("NomSousSection") = rstPiece.Fields("SousSection")
570           rstImpListePiece.Fields("IDSection") = rstPiece.Fields("IDSection")
          
575           Call rstImpListePiece.Update
        
580           iCompteurPiece = iCompteurPiece + 1
585         End If

590         Call rstImpListePiece.Close
595       End If
 
          'prochaine enreg
600       Call rstPiece.MoveNext
605     Loop

610     Call rstPiece.Close
615     Set rstPiece = Nothing
   
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        ' rapport liste piece, met dans l'ordre de ligne '
        ''''''''''''''''''''''''''''''''''''''''''''''''''
620     Call rstImpListePiece.Open("SELECT * FROM GRB_impression_Listepiece WHERE IDSoumission = '" & sNoSoumission & "' ORDER BY NoLigne", g_connData, adOpenDynamic, adLockOptimistic)
     
625     Set DR_Liste_piece.DataSource = rstImpListePiece
        
630     Call TraduireImpressionListePiece
        
        'Affiche la date
635     DR_Liste_piece.Sections("section3").Controls("lbldate").Caption = ConvertDate(Date)

        'Affiche l'entête
640     DR_Liste_piece.Sections("section4").Controls("lblsoumission").Caption = sNoSoumission
   
645     If m_eType = TYPE_PROJET Then
650       DR_Liste_piece.Sections("section4").Controls("lblprojet").Caption = rstProjSoum.Fields("IDProjet")
655     Else
660       DR_Liste_piece.Sections("section4").Controls("lblProjet").Caption = vbNullString
665     End If
    
670     DR_Liste_piece.Sections("section4").Controls("lbldescription").Caption = rstProjSoum.Fields("Description")
  
675     Call rstTemp.Open("SELECT NomClient FROM GRB_Client WHERE IDClient = " & rstProjSoum.Fields("IDClient"), g_connData, adOpenDynamic, adLockOptimistic)
   
680     DR_Liste_piece.Sections("section4").Controls("lblclient").Caption = rstTemp.Fields("NomClient")
  
685     Call rstTemp.Close
           
690     Call rstTemp.Open("SELECT NomContact FROM GRB_Contact WHERE IDContact = " & rstProjSoum.Fields("IDContact"), g_connData, adOpenDynamic, adLockOptimistic)
      
695     DR_Liste_piece.Sections("Section4").Controls("lblcontact").Caption = rstTemp.Fields("NomContact")
      
700     Call rstTemp.Close
    
        'Affiche le rapport liste des pieces
705     DR_Liste_piece.Orientation = rptOrientPortrait
  
710     Call DR_Liste_piece.Show(vbModal)
        
715     Call rstImpListePiece.Close
720     Set rstImpListePiece = Nothing

725     Set rstTemp = Nothing

730     Screen.MousePointer = vbDefault

735     Exit Sub

AfficherErreur:

740     woups "frmProjSoumMec", "ImprimerListePieces", Err, Erl
End Sub

Private Sub CalculerTempsReelsImpression(ByVal sNoProjet As String)

5       On Error GoTo AfficherErreur

10      Dim rstTotal        As ADODB.Recordset
15      Dim sDateDebut      As String
20      Dim sDateFin        As String
25      Dim sTotal          As String
30      Dim sFilterNoProjet As String

35      If Right$(sNoProjet, 2) = "99" Then
40        sFilterNoProjet = "LEFT(NoProjet, 6) = '" & Left$(sNoProjet, 6) & "'"
45      Else
50        sFilterNoProjet = "NoProjet = '" & sNoProjet & "'"
55      End If

60      Set rstTotal = New ADODB.Recordset
  
65      sDateDebut = "TIMESERIAL(Left(GRB_Punch.HeureDébut,2),RIGHT(GRB_Punch.HeureDébut,2),0)"
  
70      sDateFin = "TIMESERIAL(Left(GRB_Punch.HeureFin,2),RIGHT(GRB_Punch.HeureFin,2),0)"
  
75      sTotal = "(SUM(" & sDateFin & " - " & sDateDebut & ")* 24) As Total"

80      Call rstTotal.Open("SELECT Type, " & sTotal & " FROM GRB_Punch WHERE " & sFilterNoProjet & " And HeureFin is not null AND HeureDébut is not null GROUP BY Type", g_connData, adOpenDynamic, adLockOptimistic)

85      DR_SoumissionMec.Sections("Section5").Controls("lblTempsDessinReel").Caption = "0"
90      DR_SoumissionMec.Sections("Section5").Controls("lblTempsCoupeReel").Caption = "0"
95      DR_SoumissionMec.Sections("Section5").Controls("lblTempsMachinageReel").Caption = "0"
100     DR_SoumissionMec.Sections("Section5").Controls("lblTempsSoudureReel").Caption = "0"
105     DR_SoumissionMec.Sections("Section5").Controls("lblTempsAssemblageReel").Caption = "0"
110     DR_SoumissionMec.Sections("Section5").Controls("lblTempsPeintureReel").Caption = "0"
115     DR_SoumissionMec.Sections("Section5").Controls("lblTempsTestReel").Caption = "0"
120     DR_SoumissionMec.Sections("Section5").Controls("lblTempsInstallationReel").Caption = "0"
125     DR_SoumissionMec.Sections("Section5").Controls("lblTempsFormationReel").Caption = "0"
130     DR_SoumissionMec.Sections("Section5").Controls("lblTempsGestionReel").Caption = "0"
135     DR_SoumissionMec.Sections("Section5").Controls("lblTempsShippingReel").Caption = "0"

140     Do While Not rstTotal.EOF
145       If Not IsNull(rstTotal.Fields("Total")) Then
150         Select Case rstTotal.Fields("Type")
              Case "Dessin":       DR_SoumissionMec.Sections("Section5").Controls("lblTempsDessinReel").Caption = Round(rstTotal.Fields("Total"), 2)
155           Case "Coupe":        DR_SoumissionMec.Sections("Section5").Controls("lblTempsCoupeReel").Caption = Round(rstTotal.Fields("Total"), 2)
160           Case "Machinage":    DR_SoumissionMec.Sections("Section5").Controls("lblTempsMachinageReel").Caption = Round(rstTotal.Fields("Total"), 2)
165           Case "Soudure":      DR_SoumissionMec.Sections("Section5").Controls("lblTempsSoudureReel").Caption = Round(rstTotal.Fields("Total"), 2)
170           Case "Assemblage":   DR_SoumissionMec.Sections("Section5").Controls("lblTempsAssemblageReel").Caption = Round(rstTotal.Fields("Total"), 2)
175           Case "Peinture":     DR_SoumissionMec.Sections("Section5").Controls("lblTempsPeintureReel").Caption = Round(rstTotal.Fields("Total"), 2)
180           Case "Test":         DR_SoumissionMec.Sections("Section5").Controls("lblTempsTestReel").Caption = Round(rstTotal.Fields("Total"), 2)
185           Case "Installation": DR_SoumissionMec.Sections("Section5").Controls("lblTempsInstallationReel").Caption = Round(rstTotal.Fields("Total"), 2)
190           Case "Formation":    DR_SoumissionMec.Sections("Section5").Controls("lblTempsFormationReel").Caption = Round(rstTotal.Fields("Total"), 2)
195           Case "Gestion":      DR_SoumissionMec.Sections("Section5").Controls("lblTempsGestionReel").Caption = Round(rstTotal.Fields("Total"), 2)
200           Case "Shipping":     DR_SoumissionMec.Sections("Section5").Controls("lblTempsShippingReel").Caption = Round(rstTotal.Fields("Total"), 2)
205         End Select
210       End If

215       Call rstTotal.MoveNext
220     Loop

225     Call rstTotal.Close
  
230     Call rstTotal.Open("SELECT " & sTotal & " FROM GRB_Punch WHERE " & sFilterNoProjet & " And HeureFin is not null AND HeureDébut is not null", g_connData, adOpenDynamic, adLockOptimistic)

235     If Not IsNull(rstTotal.Fields("Total")) Then
240       DR_SoumissionMec.Sections("Section5").Controls("lblTotalTempsRHReel").Caption = Round(rstTotal.Fields("Total"), 2)
245     Else
250       DR_SoumissionMec.Sections("Section5").Controls("lblTotalTempsRHReel").Caption = "0"
255     End If

260     Call rstTotal.Close
265     Set rstTotal = Nothing

270     Exit Sub

AfficherErreur:

275     woups "frmProjSoumMec", "CalculerTempsReels", Err, Erl
End Sub

Private Sub TraduireImpressionListePiece()

5       On Error GoTo AfficherErreur

10      If m_eLangage = ANGLAIS Then
15        DR_Liste_piece.Sections("Section4").Controls("lblTitreProjet").Caption = "Project:"
20        DR_Liste_piece.Sections("Section4").Controls("lblTitreSoumission").Caption = "Quote:"
            
25        DR_Liste_piece.Sections("Section2").Controls("lblTitreQuantite").Caption = "Qty"
30        DR_Liste_piece.Sections("Section2").Controls("lblTitreNoItem").Caption = "Item No."
35        DR_Liste_piece.Sections("Section2").Controls("lblTitreManufacturier").Caption = "Manufacturer"
40        DR_Liste_piece.Sections("Section2").Controls("lblTitreID").Caption = "ID #"
          
45        DR_Liste_piece.Sections("Section3").Controls("lblNoPage").Caption = "Page %p of %P"
50      Else
55        DR_Liste_piece.Sections("Section4").Controls("lblTitreProjet").Caption = "Projet:"
60        DR_Liste_piece.Sections("Section4").Controls("lblTitreSoumission").Caption = "Soumission:"
            
65        DR_Liste_piece.Sections("Section2").Controls("lblTitreQuantite").Caption = "Qté"
70        DR_Liste_piece.Sections("Section2").Controls("lblTitreNoItem").Caption = "No. Item"
75        DR_Liste_piece.Sections("Section2").Controls("lblTitreManufacturier").Caption = "Manufacturier"
80        DR_Liste_piece.Sections("Section2").Controls("lblTitreID").Caption = "# ID"
          
85        DR_Liste_piece.Sections("Section3").Controls("lblNoPage").Caption = "Page %p de %P"
90      End If

95      Exit Sub

AfficherErreur:

100     woups "frmProjSoumMec", "TraduireImpressionListePiece", Err, Erl
End Sub

Private Sub TraduireImpressionSoumission()

5       On Error GoTo AfficherErreur

10      If m_eLangage = ANGLAIS Then
15        If m_eType = TYPE_PROJET Then
20          DR_SoumissionMec.Caption = "Mechanical Project"
25          DR_SoumissionMec.Sections("Section2").Controls("lblGrosTitre").Caption = "Mechanical project"
30        Else
35          DR_SoumissionMec.Caption = "Mechanical Quote"
40          DR_SoumissionMec.Sections("Section2").Controls("lblGrosTitre").Caption = "Mechanical quote"
45        End If
      
50        DR_SoumissionMec.Sections("Section2").Controls("lblTitreProjet").Caption = "Project:"
55        DR_SoumissionMec.Sections("Section2").Controls("lblTitreSoumission").Caption = "Quote:"
60        DR_SoumissionMec.Sections("Section2").Controls("lblTitreClient").Caption = "Client:"
65        DR_SoumissionMec.Sections("Section2").Controls("lblTitreContact").Caption = "Contact:"

        '************************************************************************************************
        'SECTION MODIFIÉ PAR GAÉTAN GINGRAS LE 07 FÉVRIER 2010
        '************************************************************************************************
      
70        DR_SoumissionMec.Sections("Section2").Controls("lblTitreQuantite").Caption = "Qty"
75        DR_SoumissionMec.Sections("Section2").Controls("lblTitreNoItem").Caption = "Item No."
80        DR_SoumissionMec.Sections("Section2").Controls("lblTitreDescription").Caption = "Description"
85        DR_SoumissionMec.Sections("Section2").Controls("lblTitreManufacturier").Caption = "Manufacturer"
90        'DR_SoumissionMec.Sections("Section2").Controls("lblTitrePrixListe").Caption = "Listed price"
95        'DR_SoumissionMec.Sections("Section2").Controls("lblTitreEscompte").Caption = "Discount"
100       DR_SoumissionMec.Sections("Section2").Controls("lblTitreCoutant").Caption = "Cost"
105       DR_SoumissionMec.Sections("Section2").Controls("lblTitreFournisseur").Caption = "Supplier"
110       DR_SoumissionMec.Sections("Section2").Controls("lblTitreTotal").Caption = "Total"

        'AJOUT PAR GAÉTAN GINGRAS 07 FÉVRIER 2010
        '****************************************************************************************
        DR_SoumissionMec.Sections("Section2").Controls("lbl_DateCommande").Caption = "Order Date"
        DR_SoumissionMec.Sections("Section2").Controls("lbl_DateReception").Caption = "Reception Date"
        '****************************************************************************************
      
115       DR_SoumissionMec.Sections("Section5").Controls("lblTitreMachinage").Caption = "Machining :"
120       DR_SoumissionMec.Sections("Section5").Controls("lblTitreCoupePreparation").Caption = "Cutting and preparing :"
125       DR_SoumissionMec.Sections("Section5").Controls("lblTitreAssemblageSoudure").Caption = "Cutting, welding and grinding :"
130       DR_SoumissionMec.Sections("Section5").Controls("lblTitreAssemblageSysteme").Caption = "System assembling :"
135       DR_SoumissionMec.Sections("Section5").Controls("lblTitrePeintureFinition").Caption = "Painting and finishing :"
140       DR_SoumissionMec.Sections("Section5").Controls("lblTitreTestFinal").Caption = "Final test :"
145       DR_SoumissionMec.Sections("Section5").Controls("lblTitreGestion").Caption = "Project management :"
150       DR_SoumissionMec.Sections("Section5").Controls("lblTitreShipping").Caption = "Shipping :"
155       DR_SoumissionMec.Sections("Section5").Controls("lblTitreConceptionDessin").Caption = "Conception and drafting :"
160       DR_SoumissionMec.Sections("Section5").Controls("lblTitreFormation").Caption = "Training :"
165       DR_SoumissionMec.Sections("Section5").Controls("lblTitreInstallation").Caption = "Installation :"
170       DR_SoumissionMec.Sections("Section5").Controls("lblTitreNbrePersonne").Caption = "Number of persons :"
175       DR_SoumissionMec.Sections("Section5").Controls("lblTitreHebergement1").Caption = "Lodging (1 bed) :"
180       DR_SoumissionMec.Sections("Section5").Controls("lblTitreHebergement2").Caption = "Lodging (2 beds) :"
185       DR_SoumissionMec.Sections("Section5").Controls("lblTitreRepas").Caption = "Meals :"
190       DR_SoumissionMec.Sections("Section5").Controls("lblTitreTransportDeplacement").Caption = "Workers freight :"
195       DR_SoumissionMec.Sections("Section5").Controls("lblTitreTransportUniteMobile").Caption = "Mobile unit freight :"
200       DR_SoumissionMec.Sections("Section5").Controls("lblTitreTransportEmballage").Caption = "Freight / Packing :"
205       DR_SoumissionMec.Sections("Section5").Controls("lblTitreTauxHoraire").Caption = "Rate / Hours"
210       DR_SoumissionMec.Sections("Section5").Controls("lblTitreTemps").Caption = "Time"
215       DR_SoumissionMec.Sections("Section5").Controls("lblHeure1").Caption = "Hours"
220       DR_SoumissionMec.Sections("Section5").Controls("lblHeure2").Caption = "Hours"
225       DR_SoumissionMec.Sections("Section5").Controls("lblHeure3").Caption = "Hours"
230       DR_SoumissionMec.Sections("Section5").Controls("lblHeure4").Caption = "Hours"
235       DR_SoumissionMec.Sections("Section5").Controls("lblHeure5").Caption = "Hours"
240       DR_SoumissionMec.Sections("Section5").Controls("lblHeure6").Caption = "Hours"
245       DR_SoumissionMec.Sections("Section5").Controls("lblHeure7").Caption = "Hours"
250       DR_SoumissionMec.Sections("Section5").Controls("lblHeure8").Caption = "Hours"
255       DR_SoumissionMec.Sections("Section5").Controls("lblHeure9").Caption = "Hours"
260       DR_SoumissionMec.Sections("Section5").Controls("lblHeure10").Caption = "Hours"
265       DR_SoumissionMec.Sections("Section5").Controls("lblHeure11").Caption = "Hours"
270       DR_SoumissionMec.Sections("Section5").Controls("lblHeure12").Caption = "Hours"
285       DR_SoumissionMec.Sections("Section5").Controls("lblJour1").Caption = "Days"
290       DR_SoumissionMec.Sections("Section5").Controls("lblJour2").Caption = "Days"
295       DR_SoumissionMec.Sections("Section5").Controls("lblTitreTotalTemps").Caption = "Time total:"
300       DR_SoumissionMec.Sections("Section5").Controls("lblTitreTotalPiece").Caption = "Parts total:"
305       DR_SoumissionMec.Sections("Section5").Controls("lblTitreImprevue").Caption = "Unforeseen:"
310       DR_SoumissionMec.Sections("Section5").Controls("lblTitreAutre").Caption = "Other:"
315       DR_SoumissionMec.Sections("Section5").Controls("lblTitreGrandTotal").Caption = "Final price:"
      
320       DR_SoumissionMec.Sections("Section3").Controls("lblNoPage").Caption = "Page %p of %P"
325       DR_SoumissionMec.Sections("Section5").Controls("lblTitrePrixReception").Caption = "Receiving up to date"
330       DR_SoumissionMec.Sections("Section5").Controls("lblTitrePrixSoumission").Caption = "Quote Price"
335       DR_SoumissionMec.Sections("Section5").Controls("lblTitreForfait").Caption = "Package Deal"
340     Else
345       If m_eType = TYPE_PROJET Then
350         DR_SoumissionMec.Caption = "Projet Mécanique"
355         DR_SoumissionMec.Sections("Section2").Controls("lblGrosTitre").Caption = "Projet mécanique"
360       Else
365         DR_SoumissionMec.Caption = "Soumission Mécanique"
370         DR_SoumissionMec.Sections("Section2").Controls("lblGrosTitre").Caption = "Soumission mécanique"
375       End If
     
380       DR_SoumissionMec.Sections("Section2").Controls("lblTitreProjet").Caption = "Projet:"
385       DR_SoumissionMec.Sections("Section2").Controls("lblTitreSoumission").Caption = "Soumission:"
390       DR_SoumissionMec.Sections("Section2").Controls("lblTitreClient").Caption = "Client:"
395       DR_SoumissionMec.Sections("Section2").Controls("lblTitreContact").Caption = "Contact:"
     
400       DR_SoumissionMec.Sections("Section2").Controls("lblTitreQuantite").Caption = "Qté"
405       DR_SoumissionMec.Sections("Section2").Controls("lblTitreNoItem").Caption = "No. item"
410       DR_SoumissionMec.Sections("Section2").Controls("lblTitreDescription").Caption = "Description"
415       DR_SoumissionMec.Sections("Section2").Controls("lblTitreManufacturier").Caption = "Manufacturier"
420       'DR_SoumissionMec.Sections("Section2").Controls("lblTitrePrixListe").Caption = "Prix de liste"
425       'DR_SoumissionMec.Sections("Section2").Controls("lblTitreEscompte").Caption = "Escompte"
430       DR_SoumissionMec.Sections("Section2").Controls("lblTitreCoutant").Caption = "Coûtant"
435       DR_SoumissionMec.Sections("Section2").Controls("lblTitreFournisseur").Caption = "Fournisseur"
440       DR_SoumissionMec.Sections("Section2").Controls("lblTitreTotal").Caption = "Total"

        'AJOUT PAR GAÉTAN GINGRAS 07 FÉVRIER 2010
        '****************************************************************************************
        DR_SoumissionMec.Sections("Section2").Controls("lbl_DateCommande").Caption = "Date commandé"
        DR_SoumissionMec.Sections("Section2").Controls("lbl_DateReception").Caption = "Date reçu"
        '****************************************************************************************

        '************************************************************************************************
        'FIN DE LA SECTION DE MODIFICATION
        '************************************************************************************************
    
445       DR_SoumissionMec.Sections("Section5").Controls("lblTitreMachinage").Caption = "Machinage :"
450       DR_SoumissionMec.Sections("Section5").Controls("lblTitreCoupePreparation").Caption = "Coupe et préparation :"
455       DR_SoumissionMec.Sections("Section5").Controls("lblTitreAssemblageSoudure").Caption = "Soudure et meulage :"
460       DR_SoumissionMec.Sections("Section5").Controls("lblTitreAssemblageSysteme").Caption = "Assemblage des systèmes :"
465       DR_SoumissionMec.Sections("Section5").Controls("lblTitrePeintureFinition").Caption = "Peinture et finition :"
470       DR_SoumissionMec.Sections("Section5").Controls("lblTitreTestFinal").Caption = "Tests finaux :"
475       DR_SoumissionMec.Sections("Section5").Controls("lblTitreGestion").Caption = "Gestion du projet :"
480       DR_SoumissionMec.Sections("Section5").Controls("lblTitreShipping").Caption = "Expédition :"
485       DR_SoumissionMec.Sections("Section5").Controls("lblTitreConceptionDessin").Caption = "Conception et dessin :"
490       DR_SoumissionMec.Sections("Section5").Controls("lblTitreFormation").Caption = "Formation du personnel :"
495       DR_SoumissionMec.Sections("Section5").Controls("lblTitreInstallation").Caption = "Installation :"
500       DR_SoumissionMec.Sections("Section5").Controls("lblTitreNbrePersonne").Caption = "Nombre de personnes :"
505       DR_SoumissionMec.Sections("Section5").Controls("lblTitreHebergement1").Caption = "Hébergement (1 lit):"
510       DR_SoumissionMec.Sections("Section5").Controls("lblTitreHebergement2").Caption = "Hébergement (2 lits) :"
515       DR_SoumissionMec.Sections("Section5").Controls("lblTitreRepas").Caption = "Repas :"
520       DR_SoumissionMec.Sections("Section5").Controls("lblTitreTransportDeplacement").Caption = "Transport / Déplacement :"
525       DR_SoumissionMec.Sections("Section5").Controls("lblTitreTransportUniteMobile").Caption = "Transport de l'unité mobile :"
530       DR_SoumissionMec.Sections("Section5").Controls("lblTitreTransportEmballage").Caption = "Transport / Emballage :"
535       DR_SoumissionMec.Sections("Section5").Controls("lblTitreTauxHoraire").Caption = "Taux horaire"
540       DR_SoumissionMec.Sections("Section5").Controls("lblTitreTemps").Caption = "Temps"
545       DR_SoumissionMec.Sections("Section5").Controls("lblHeure1").Caption = "Heures"
550       DR_SoumissionMec.Sections("Section5").Controls("lblHeure2").Caption = "Heures"
555       DR_SoumissionMec.Sections("Section5").Controls("lblHeure3").Caption = "Heures"
560       DR_SoumissionMec.Sections("Section5").Controls("lblHeure4").Caption = "Heures"
565       DR_SoumissionMec.Sections("Section5").Controls("lblHeure5").Caption = "Heures"
570       DR_SoumissionMec.Sections("Section5").Controls("lblHeure6").Caption = "Heures"
575       DR_SoumissionMec.Sections("Section5").Controls("lblHeure7").Caption = "Heures"
580       DR_SoumissionMec.Sections("Section5").Controls("lblHeure8").Caption = "Heures"
585       DR_SoumissionMec.Sections("Section5").Controls("lblHeure9").Caption = "Heures"
590       DR_SoumissionMec.Sections("Section5").Controls("lblHeure10").Caption = "Heures"
595       DR_SoumissionMec.Sections("Section5").Controls("lblHeure11").Caption = "Heures"
600       DR_SoumissionMec.Sections("Section5").Controls("lblHeure12").Caption = "Heures"
615       DR_SoumissionMec.Sections("Section5").Controls("lblJour1").Caption = "Jours"
620       DR_SoumissionMec.Sections("Section5").Controls("lblJour2").Caption = "Jours"
625       DR_SoumissionMec.Sections("Section5").Controls("lblTitreTotalTemps").Caption = "Temps total:"
630       DR_SoumissionMec.Sections("Section5").Controls("lblTitreTotalPiece").Caption = "Total des pièces:"
635       DR_SoumissionMec.Sections("Section5").Controls("lblTitreImprevue").Caption = "Imprévue:"
640       DR_SoumissionMec.Sections("Section5").Controls("lblTitreAutre").Caption = "Autre:"
645       DR_SoumissionMec.Sections("Section5").Controls("lblTitreGrandTotal").Caption = "Grand total:"
    
650       DR_SoumissionMec.Sections("Section3").Controls("lblNoPage").Caption = "Page %p de %P"
655       DR_SoumissionMec.Sections("Section5").Controls("lblTitrePrixReception").Caption = "Réception jusqu'à maintenant"
660       DR_SoumissionMec.Sections("Section5").Controls("lblTitrePrixSoumission").Caption = "$ Soumission"
665       DR_SoumissionMec.Sections("Section5").Controls("lblTitreForfait").Caption = "Forfait"
670     End If

675     Exit Sub

AfficherErreur:

680     woups "frmProjSoumMec", "TraduireImpressionSoumission", Err, Erl
End Sub

Private Sub cmdModifier_Click()

5       On Error GoTo AfficherErreur

10      Dim rstProjSoum As ADODB.Recordset
15      Dim sUser       As String

20      Set m_collQteSupp = New Collection
25      Set m_collDateSupp = New Collection
30      Set m_collHeureSupp = New Collection
35      Set m_collNoItemSupp = New Collection

40      If Right$(txtNoProjSoum.Text, 2) = "99" Then
45        If m_eType = TYPE_PROJET Then
50          Call MsgBox("Ce projet ne peut pas être modifié!", vbOKOnly, "Erreur")
55        Else
60          Call MsgBox("Cette soumission ne peut pas être modifiée!", vbOKOnly, "Erreur")
65        End If
        
70        Exit Sub
75      End If

80      Set rstProjSoum = New ADODB.Recordset

85      Call rstProjSoum.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

90      If rstProjSoum.Fields("Ouvert") = False Or rstProjSoum.Fields("Verrouillé") = True Then
95        If rstProjSoum.Fields("Ouvert") = False Then
100         If m_eType = TYPE_PROJET Then
105           Call MsgBox("Ce projet est fermé!", vbOKOnly, "Erreur")
110         Else
115           Call MsgBox("Cette soumission est fermée!", vbOKOnly, "Erreur")
120         End If
125       Else
130         If m_eType = TYPE_PROJET Then
135           Call MsgBox("Ce projet est verrouillé!", vbOKOnly, "Erreur")
140         Else
145           Call MsgBox("Cette soumission est verrouillée!", vbOKOnly, "Erreur")
150         End If
155       End If

160       Call rstProjSoum.Close
165       Set rstProjSoum = Nothing

170       Exit Sub
175     End If

180     Call rstProjSoum.Close
185     Set rstProjSoum = Nothing

190     If VerifierSiOuvert(sUser) = False Then
          'Modifier une soumission
195       If cmbProjSoum.ListIndex > -1 Then
200         If m_eType = TYPE_SOUMISSION Then
205           If VerifierSiDejaProjet = True Then
210             Call MsgBox("Vous ne pouvez pas modifier cette soumission, le projet a déjà été crée!", vbOKOnly, "Erreur")
        
215             Exit Sub
220           End If
225         End If
  
230         Screen.MousePointer = vbHourglass
  
            'Débarre les champs
235         Call BarrerChamps(False)
    
            'Pour pouvoir afficher le dernier enregistrement affiché quand la personne va
            'enregistrer ou annuler
240         m_sAncienProjSoum = txtNoProjSoum.Text
  
245         m_bModeAjout = False
250         m_bModeAffichage = False
          
            'Rapetisse le listview de la soumission pour afficher le lvwPiece
255         lvwSoumission.Height = lvwSoumission.Height * 0.49
260         lvwSoumission.Top = lvwPieces.Top + lvwPieces.Height + (lvwSoumission.Height * 0.02)
           
265         Call RemplirProjSoum
           
270         Call AfficherControles(MODE_AJOUT_MODIF)
    
275         Call UpdateOrdre
    
280         Call CalculerPrix
  
285         Call OuvrirProjSoum(True)
  
290         Screen.MousePointer = vbDefault
295       End If
300     Else
305       If m_eType = TYPE_PROJET Then
310         Call MsgBox("Ce projet est en modification par " & sUser & "!", vbOKOnly, "Erreur")
315       Else
320         Call MsgBox("Cette soumission est en modification par " & sUser & "!", vbOKOnly, "Erreur")
325       End If
330     End If

335     Exit Sub

AfficherErreur:

340     woups "frmProjSoumMec", "cmdModifier_Click", Err, Erl
End Sub

Private Sub InitialiserNouveauxTaux()

5       On Error GoTo AfficherErreur

10      Dim rstConfig As ADODB.Recordset
  
15      Set rstConfig = New ADODB.Recordset
  
20      Call rstConfig.Open("SELECT TauxDessinMec, TauxCoupe, TauxMachinage, TauxSoudure, TauxAssemblageMec, TauxPeinture, TauxTestMec, TauxFormationMec, TauxInstallationMec, TauxGestionProjetsMec, TauxShippingMec, Hebergement1, Hebergement2, Repas, Standard, UniteMobile FROM GRB_Config", g_connData, adOpenDynamic, adLockOptimistic)
  
25      If Not IsNull(rstConfig.Fields("TauxDessinMec")) Then
30        m_sTauxDessin = rstConfig.Fields("TauxDessinMec")
35      Else
40        m_sTauxDessin = "0"
45      End If

50      If Not IsNull(rstConfig.Fields("TauxCoupe")) Then
55        m_sTauxCoupe = rstConfig.Fields("TauxCoupe")
60      Else
65        m_sTauxCoupe = "0"
70      End If

75      If Not IsNull(rstConfig.Fields("TauxMachinage")) Then
80        m_sTauxMachinage = rstConfig.Fields("TauxMachinage")
85      Else
90        m_sTauxMachinage = "0"
95      End If

100     If Not IsNull(rstConfig.Fields("TauxSoudure")) Then
105       m_sTauxSoudure = rstConfig.Fields("TauxSoudure")
110     Else
115       m_sTauxSoudure = "0"
120     End If

125     If Not IsNull(rstConfig.Fields("TauxAssemblageMec")) Then
130       m_sTauxAssemblage = rstConfig.Fields("TauxAssemblageMec")
135     Else
140       m_sTauxAssemblage = "0"
145     End If

150     If Not IsNull(rstConfig.Fields("TauxPeinture")) Then
155       m_sTauxPeinture = rstConfig.Fields("TauxPeinture")
160     Else
165       m_sTauxPeinture = "0"
170     End If

175     If Not IsNull(rstConfig.Fields("TauxTestMec")) Then
180       m_sTauxTest = rstConfig.Fields("TauxTestMec")
185     Else
190       m_sTauxTest = "0"
195     End If

200     If Not IsNull(rstConfig.Fields("TauxInstallationMec")) Then
205       m_sTauxInstallation = rstConfig.Fields("TauxInstallationMec")
210     Else
215       m_sTauxInstallation = "0"
220     End If

225     If Not IsNull(rstConfig.Fields("TauxFormationMec")) Then
230       m_sTauxFormation = rstConfig.Fields("TauxFormationMec")
235     Else
240       m_sTauxFormation = "0"
245     End If

250     If Not IsNull(rstConfig.Fields("TauxGestionProjetsMec")) Then
255       m_sTauxGestion = rstConfig.Fields("TauxGestionProjetsMec")
260     Else
265       m_sTauxGestion = "0"
270     End If

275     If Not IsNull(rstConfig.Fields("TauxShippingMec")) Then
280       m_sTauxShipping = rstConfig.Fields("TauxShippingMec")
285     Else
290       m_sTauxShipping = "0"
295     End If

300     If m_eType = TYPE_PROJET Then
305       m_sTauxHebergement1 = "0"
310       m_sTauxHebergement2 = "0"
315       m_sTauxRepas = "0"
320       m_sTauxTransport = "0"
325       m_sTauxUniteMobile = "0"
330     Else
335       m_sTauxHebergement1 = rstConfig.Fields("Hebergement1")
340       m_sTauxHebergement2 = rstConfig.Fields("Hebergement2")
345       m_sTauxRepas = rstConfig.Fields("Repas")
350       m_sTauxTransport = rstConfig.Fields("Standard")
355       m_sTauxUniteMobile = rstConfig.Fields("UniteMobile")
360     End If
    
365     Call rstConfig.Close
370     Set rstConfig = Nothing

375     Exit Sub

AfficherErreur:

380     woups "frmProjSoumMec", "InitialiserNouveauxTaux", Err, Erl
End Sub

Private Sub cmdsupprimer_Click()

5       On Error GoTo AfficherErreur

10      Dim rstProjSoum As ADODB.Recordset
15      Dim rstProjet   As ADODB.Recordset
20      Dim iReponse    As Integer
25      Dim sSoumission As String
30      Dim sUser       As String
35      Dim iExtension  As Integer

        'S'il y a des enregistrements
40      If cmbProjSoum.ListCount > 0 Then
45        If Right$(txtNoProjSoum.Text, 2) = "99" Then
50          If m_eType = TYPE_PROJET Then
55            Call MsgBox("Vous ne pouvez pas supprimer ce projet!", vbOKOnly, "Erreur")
60          Else
65            Call MsgBox("Vous ne pouvez pas supprimer cette soumission!", vbOKOnly, "Erreur")
70          End If

75          Exit Sub
80        End If

85        Set rstProjSoum = New ADODB.Recordset

90        Call rstProjSoum.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

95        If rstProjSoum.Fields("Ouvert") = False Or rstProjSoum.Fields("Verrouillé") = True Then
100         If rstProjSoum.Fields("Ouvert") = False Then
105           If m_eType = TYPE_PROJET Then
110             Call MsgBox("Ce projet est fermé!", vbOKOnly, "Erreur")
115           Else
120             Call MsgBox("Cette soumission est fermée!", vbOKOnly, "Erreur")
125           End If
130         Else
135           If m_eType = TYPE_PROJET Then
140             Call MsgBox("Ce projet est verrouillé!", vbOKOnly, "Erreur")
145           Else
150             Call MsgBox("Cette soumission est verrouillée!", vbOKOnly, "Erreur")
155           End If
160         End If

165         Call rstProjSoum.Close
170         Set rstProjSoum = Nothing

175         Exit Sub
180       End If

185       Call rstProjSoum.Close
190       Set rstProjSoum = Nothing
  
195       If m_eType = TYPE_SOUMISSION Then
200         If VerifierSiDejaProjet = True Then
205           Call MsgBox("Vous ne pouvez pas supprimer cette soumission, le projet a déjà été crée!", vbOKOnly, "Erreur")
       
210           Exit Sub
215         End If
220       End If

225       If VerifierSiOuvert(sUser) = False Then
            'Valider le choix
230         If m_eType = TYPE_PROJET Then
235           iReponse = MsgBox("Voulez-vous vraiment EFFACER LE PROJET " & txtNoProjSoum.Text & "?", vbYesNo)

240           If iReponse = vbYes Then
245             Call frmValiderSuppression.Afficher(True, txtNoProjSoum.Text, Me)

250             If m_bValide = True Then
255               iReponse = vbYes
260             Else
265               iReponse = vbNo
270             End If
275           End If
280         Else
285           iReponse = MsgBox("Voulez-vous vraiment EFFACER LA SOUMISSION " & txtNoProjSoum.Text & "?", vbYesNo)

290           If iReponse = vbYes Then
295             Call frmValiderSuppression.Afficher(False, txtNoProjSoum.Text, Me)

300             If m_bValide = True Then
305               iReponse = vbYes
310             Else
315               iReponse = vbNo
320             End If
325           End If
330         End If
          
            'S'il veut vraiment effacer
335         If iReponse = vbYes Then
              'Si c'est un projet
340           If m_eType = TYPE_PROJET Then
345             Set rstProjet = New ADODB.Recordset

350             Call rstProjet.Open("SELECT IDSoumission FROM GRB_ProjetMec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

355             If Not IsNull(rstProjet.Fields("IDSoumission")) Then
360               sSoumission = rstProjet.Fields("IDSoumission")
365             Else
370               sSoumission = vbNullString
375             End If

380             Call rstProjet.Close
385             Set rstProjet = Nothing

                'Efface les Pièces
390             Call g_connData.Execute("DELETE * FROM GRB_Projet_Pieces WHERE IDProjet = '" & txtNoProjSoum.Text & "' AND Type = 'M'")
                
395             If IsNumeric(Right$(txtNoProjSoum.Text, 2)) Then
400               iExtension = CInt(Right$(txtNoProjSoum.Text, 2))
405             Else
410               iExtension = 0
415             End If

420             If (iExtension >= 60 And iExtension <= 79) Or (iExtension >= 80 And iExtension <= 98) Then
425               Set rstProjSoum = New ADODB.Recordset

430               Call rstProjSoum.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

435               Call g_connData.Execute("DELETE * FROM GRB_Projet_Pieces WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & rstProjSoum.Fields("LiaisonChargeable") & "' AND Provenance = '" & iExtension & "'")

440               Call CalculerTotalRecordset(Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & rstProjSoum.Fields("LiaisonChargeable"))
                  
445               Call rstProjSoum.Close
450               Set rstProjSoum = Nothing
455             End If
                
                'Efface les modifications
460             Call g_connData.Execute("DELETE * FROM GRB_Projet_Modif WHERE IDProjet = '" & txtNoProjSoum.Text & "' AND Type = 'M'")
                
                'Efface le projet
465             Call g_connData.Execute("DELETE * FROM GRB_ProjetMec WHERE IDProjet = '" & txtNoProjSoum.Text & "'")
                
                'Efface le projet dans la table GRB_ProjSoum
470             Call g_connData.Execute("DELETE * FROM GRB_ProjSoum WHERE IDProjSoum = '" & txtNoProjSoum.Text & "'")

475             Set rstProjSoum = New ADODB.Recordset

480             Call rstProjSoum.Open("SELECT Ouvert FROM GRB_ProjSoum WHERE IDProjSoum = '" & sSoumission & "'", g_connData, adOpenDynamic, adLockOptimistic)

485             If Not rstProjSoum.EOF Then
490               rstProjSoum.Fields("Ouvert") = True

495               Call rstProjSoum.Update
500             End If

505             Call rstProjSoum.Close
510             Set rstProjSoum = Nothing
515           Else
                'Efface les pièces
520             Call g_connData.Execute("DELETE * FROM GRB_Soumission_Pieces WHERE IDSoumission = '" & txtNoProjSoum.Text & "' AND Type = 'M'")
          
                'Efface les modifications
525             Call g_connData.Execute("DELETE * FROM GRB_Soumission_Modif WHERE IDSoumission = '" & txtNoProjSoum.Text & "' AND Type = 'M'")
        
                'Efface la soumission
530             Call g_connData.Execute("DELETE * FROM GRB_SoumissionMec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'")
      
                'Efface la soumission dans la table GRB_ProjSoum
535             Call g_connData.Execute("DELETE * FROM GRB_ProjSoum WHERE IDProjSoum = '" & txtNoProjSoum.Text & "'")
540           End If
              
545           If m_eType = TYPE_PROJET Then
550             Call RecreerProjetCumulatif
555           Else
560             Call RecreerSoumissionCumulatif
565           End If
              
              'Affiche la premiere soumission
570           Call AfficherProjSoum(vbNullString)
575         End If
580       Else
585         If m_eType = TYPE_PROJET Then
590           Call MsgBox("Ce projet est en modification par " & sUser & "!", vbOKOnly, "Erreur")
595         Else
600           Call MsgBox("Cette soumission est en modification par " & sUser & "!", vbOKOnly, "Erreur")
605         End If
610       End If
615     End If

620     Exit Sub

AfficherErreur:

625     woups "frmProjSoumMec", "cmdSupprimer_Click", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

10      Call Unload(frmChoixProjSoum)
    
15      m_eLangage = FRANCAIS
    
20      cmdAnglaisFrancais.Caption = "Anglais"
    
        'Initialise le tri à PIECE_GRB
25      cmbTri.ListIndex = I_CMB_PIECE
     
        'Donne accès aux boutons selon le groupe
30      Call ActiverBoutonsGroupe
  
        'Rempli le combo des clients
35      Call RemplirComboClients(vbNullString)
  
        'Rempli le combo des contacts
40      Call RemplirComboSections
  
        'Rempli le combo des catégories de pièce
45      Call RemplirComboCategoriesPieces
              
50      cmbOuvertFerme.ListIndex = I_CMB_OUVERT
              
55      If m_eType = TYPE_PROJET Then
60        cmbChoix.ListIndex = I_IDX_PROJET
65      Else
70        cmbChoix.ListIndex = I_IDX_SOUMISSION
75      End If

80      Exit Sub

AfficherErreur:

85      woups "frmProjSoumMec", "Form_Load", Err, Erl
End Sub

Private Sub RemplirColonnes()

5       On Error GoTo AfficherErreur

        'Méthode pour afficher les colonnes selon le groupe de sécurité.
  
        'Ceux qui n'ont pas le droit de modifier les soumissions ou les projets, n'ont pas
        'le droit de voir les prix, donc il faut cacher les colonnes
10      Dim bModif      As Boolean
15      Dim bCacherPrix As Boolean
    
        'Si le type d'affichage est "Projet"
20      If m_eType = TYPE_PROJET Then
          'Si l'utilisateur n'a pas le droit de modification sur les projets
25        If m_bModifProj = False Then
            'On cache les prix
30          bCacherPrix = True
35        End If
40      Else
          'Si l'utilisateur n'a pas le droit de modification sur les soumissions
45        If m_bModifSoum = False Then
            'On cache les prix
50          bCacherPrix = True
55        End If
60      End If
    
        'Si on cache les prix
65      If bCacherPrix = True Then
70        m_bDroitPrix = False
    
          'Si les bonnes colonnes sont déjà toutes affichées
75        If lvwSoumission.ColumnHeaders.count = 9 Then
80          Exit Sub
85        End If
90      Else
95        m_bDroitPrix = True
    
          'Si les bonnes colonnes sont déjà toutes affichées
100       If lvwSoumission.ColumnHeaders.count = 15 Then
105         Exit Sub
110       End If
115     End If
    
        'Il faut enlever les colonnes avant d'en ajouter d'autres
120     Call lvwSoumission.ColumnHeaders.Clear
      
125     Call lvwSoumission.ColumnHeaders.Add(, , "Qté", 650.1418)
130     Call lvwSoumission.ColumnHeaders.Add(, , "No. Item", 2200)
135     Call lvwSoumission.ColumnHeaders.Add(, , "Description", 3809.746)
140     Call lvwSoumission.ColumnHeaders.Add(, , "Manufacturier", 1154.8348)
    
145     If bCacherPrix = False Then
150       Call lvwSoumission.ColumnHeaders.Add(, , "Prix listé", 920.1261, vbRightJustify)
155       Call lvwSoumission.ColumnHeaders.Add(, , "Escompte", 884.9765, vbRightJustify)
160       Call lvwSoumission.ColumnHeaders.Add(, , "Prix net", 920.1261, vbRightJustify)
165     End If
    
170     Call lvwSoumission.ColumnHeaders.Add(, , "Distributeur", 1005.1655)
      
175     If bCacherPrix = False Then
180       Call lvwSoumission.ColumnHeaders.Add(, , "TOTAL", 1099.8426, vbRightJustify)
185       Call lvwSoumission.ColumnHeaders.Add(, , "Profit", 920.1261, vbRightJustify)
190     End If

195     Call lvwSoumission.ColumnHeaders.Add(, , "Commentaire", 1000)

200     If m_eType = TYPE_PROJET Then
205       If bCacherPrix = False Then
210         Call lvwSoumission.ColumnHeaders.Add(, , "Facturation", 1440)
215       End If

220       Call lvwSoumission.ColumnHeaders.Add(, , "Date Commande", 1440)
225       Call lvwSoumission.ColumnHeaders.Add(, , "Date Requise", 1440)
230       Call lvwSoumission.ColumnHeaders.Add(, , "Commandé par", 1440)
235       Call lvwSoumission.ColumnHeaders.Add(, , "No Séquentiel", 1440)
240     End If

245     Call lvwSoumission.ColumnHeaders.Add(, , "Provenance", 1440)

250     Exit Sub

AfficherErreur:

255     woups "frmProjSoumMec", "RemplirColonnes", Err, Erl
End Sub

Private Sub BarrerChamps(ByVal bBarrer As Boolean)

5       On Error GoTo AfficherErreur

        'Méthode qui barre ou débarre les champs d'après la variable bBarrer
10      txtDescription.Locked = bBarrer
15      txtNbreManuel.Locked = bBarrer
20      txtPrixManuel.Locked = bBarrer

25      Exit Sub

AfficherErreur:

30      woups "frmProjSoumMec", "BarrerChamps", Err, Erl
End Sub

Private Sub ViderChamps()

5       On Error GoTo AfficherErreur

        'Méthode qui initialise les champs
10      txtClient.Text = vbNullString
15      txtcontact.Text = vbNullString
20      txtDescription.Text = vbNullString
25      txtNbreManuel.Text = 0
30      txtPrixManuel.Text = 0
35      txtPrixTotal.Text = 0
40      txtProfit.Text = 0
45      txtPrixReception.Text = Conversion("0", MODE_ARGENT)
50      txtPrixSoumission.Text = Conversion("0", MODE_ARGENT)
55      txtCommission.Text = 0
60      txtTotalPieces.Text = 0
65      txtTotalTemps.Text = 0
70      txtImprevus.Text = 0
75      txtNoSoumission.Text = vbNullString
80      txtCheminPhotos.Text = vbNullString
85      txtForfait.Text = vbNullString
90      lblForfaitInitiale.Caption = vbNullString
  
95      cmbclient.ListIndex = -1

100     Call lvwSoumission.ListItems.Clear

105     Exit Sub

AfficherErreur:

110     woups "frmProjSoumMec", "ViderChamps", Err, Erl
End Sub

Private Sub RemplirComboProjSoum(ByVal sNoProjSoum As String)

5       On Error GoTo AfficherErreur

        'Rempli le combo des soumissions
10      Dim rstProjSoum As ADODB.Recordset
  
        'Il faut vider le combo avant de le remplir
15      Call cmbProjSoum.Clear
  
20      Set rstProjSoum = New ADODB.Recordset
  
        'Ouvre le recordset selon le type
25      If m_eType = TYPE_PROJET Then
30        If cmbOuvertFerme.ListIndex = I_CMB_OUVERT Then
35          Call rstProjSoum.Open("SELECT IDProjet FROM GRB_ProjetMec INNER JOIN GRB_ProjSoum ON GRB_ProjetMec.IDProjet = GRB_ProjSoum.IDProjSoum WHERE Ouvert = True ORDER BY IDProjet DESC", g_connData, adOpenDynamic, adLockOptimistic)
40        Else
45          Call rstProjSoum.Open("SELECT IDProjet FROM GRB_ProjetMec ORDER BY IDProjet DESC", g_connData, adOpenDynamic, adLockOptimistic)
50        End If
55      Else
60        If cmbOuvertFerme.ListIndex = I_CMB_OUVERT Then
65          Call rstProjSoum.Open("SELECT IDSoumission FROM GRB_SoumissionMec INNER JOIN GRB_ProjSoum ON GRB_SoumissionMec.IDSoumission = GRB_ProjSoum.IDProjSoum WHERE Ouvert = True ORDER BY IDSoumission DESC", g_connData, adOpenDynamic, adLockOptimistic)
70        Else
75          Call rstProjSoum.Open("SELECT IDSoumission FROM GRB_SoumissionMec ORDER BY IDSoumission DESC", g_connData, adOpenDynamic, adLockOptimistic)
80        End If
85      End If

        'Tant que ce n'est pas la fin des enregistrements
90      Do While Not rstProjSoum.EOF
          'On met le numéro de la soumission dans le combo des soumissions
95        If m_eType = TYPE_PROJET Then
100         Call cmbProjSoum.AddItem(rstProjSoum.Fields("IDProjet"))
105       Else
110         Call cmbProjSoum.AddItem(rstProjSoum.Fields("IDSoumission"))
115       End If

120       Call rstProjSoum.MoveNext
125     Loop
     
130     Call rstProjSoum.Close
135     Set rstProjSoum = Nothing

        'Si le combo n'est pas vide, on sélectionne l'item voulu ou le premier
140     If cmbProjSoum.ListCount > 0 Then
          'Si il y a un numéro de projet
145       If sNoProjSoum <> vbNullString Then
            'On le sélectionne dans le combo
150         Call RechercherProjSoum(sNoProjSoum)
155       Else
            'Sinon, on sélectionne le premier
160         cmbProjSoum.ListIndex = 0
165       End If
170     End If

175     Exit Sub

AfficherErreur:

180     woups "frmProjSoumMec", "RemplirComboProjSoum", Err, Erl
End Sub

Private Sub CalculerPrixReel(ByVal sNoItem As String)

5       On Error GoTo AfficherErreur

10      Dim rstPieceFRS As ADODB.Recordset
15      Dim rstConfig   As ADODB.Recordset
20      Dim sPrixCalcul As String
25      Dim sTauxUSA    As String
30      Dim sTauxSPA    As String

35      Set rstConfig = New ADODB.Recordset

40      Call rstConfig.Open("SELECT TauxAmericain, TauxEspagnol FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)
  
45      sTauxUSA = rstConfig.Fields("TauxAmericain")
50      sTauxSPA = rstConfig.Fields("TauxEspagnol")

55      Call rstConfig.Close
60      Set rstConfig = Nothing

65      Set rstPieceFRS = New ADODB.Recordset
  
70      Call rstPieceFRS.Open("SELECT PrixReel, PRIX_NET, PRIX_SP, DeviseMonétaire FROM GRB_PiecesFRS WHERE PIECE = '" & Replace(sNoItem, "'", "''") & "' AND Type = 'M'", g_connData, adOpenDynamic, adLockOptimistic)
    
75      Do While Not rstPieceFRS.EOF
80        If rstPieceFRS.Fields("PRIX_NET") <> vbNullString Then
85          sPrixCalcul = Replace(rstPieceFRS.Fields("PRIX_NET"), ".", ",")
90        Else
95          sPrixCalcul = Replace(rstPieceFRS.Fields("PRIX_SP"), ".", ",")
100       End If
      
105       If rstPieceFRS.Fields("DeviseMonétaire") = "USA" Then
110         rstPieceFRS.Fields("PrixReel") = Conversion(CStr(Round(CDbl(sPrixCalcul) / CDbl(sTauxUSA), 4)), MODE_DECIMAL, 4)
115       Else
120         If rstPieceFRS.Fields("DeviseMonétaire") = "SPA" Then
125           rstPieceFRS.Fields("PrixReel") = Conversion(CStr(Round(CDbl(sPrixCalcul) / CDbl(sTauxSPA), 4)), MODE_DECIMAL, 4)
130         Else
135           rstPieceFRS.Fields("PrixReel") = Conversion(sPrixCalcul, MODE_DECIMAL, 4)
140         End If
145       End If

150       Call rstPieceFRS.Update
    
155       Call rstPieceFRS.MoveNext
160     Loop
    
165     Call rstPieceFRS.Close
170     Set rstPieceFRS = Nothing

175     Exit Sub

AfficherErreur:

180     woups "frmProjSoumMec", "CalculerPrixReel", Err, Erl
End Sub

Private Sub RemplirComboFournisseur()

5       On Error GoTo AfficherErreur

10      Dim rstFRS As ADODB.Recordset

        'Il faut vider le combo avant de le remplir
15      Call cmbfrs.Clear

20      Set rstFRS = New ADODB.Recordset

25      Call rstFRS.Open("SELECT GRB_PiecesFRS.*, GRB_Fournisseur.NomFournisseur FROM GRB_PiecesFRS INNER JOIN GRB_Fournisseur ON GRB_PiecesFRS.IDFRS = GRB_Fournisseur.IDFRS WHERE PIECE = '" & Replace(lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE), "'", "''") & "' ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)

        'Tant que ce n'est pas la fin des enregistrements
30      Do While Not rstFRS.EOF
35        Call cmbfrs.AddItem(rstFRS.Fields("NomFournisseur"))

40        cmbfrs.ItemData(cmbfrs.newIndex) = rstFRS.Fields("IDFRS")

45        Call rstFRS.MoveNext
50      Loop

55      Exit Sub

AfficherErreur:

60      woups "frmProjSoumMec", "RemplirComboFournisseur", Err, Erl
End Sub

Private Sub RemplirListViewFournisseur()

5       On Error GoTo AfficherErreur

        'Rempli le listview des distributeur pour une pièce choisie
10      Dim rstPieceFRS As ADODB.Recordset
15      Dim rstContact  As ADODB.Recordset
20      Dim rstFRS      As ADODB.Recordset
25      Dim rstInv      As ADODB.Recordset
30      Dim itmFRS      As ListItem
35      Dim iCompteur   As Integer
40      Dim iNoClient   As Integer
45      Dim bAjouterDP  As Boolean
50      Dim sDevise     As String
55      Dim lColor      As Long
  
60      Set rstPieceFRS = New ADODB.Recordset
65      Set rstContact = New ADODB.Recordset
70      Set rstFRS = New ADODB.Recordset
  
        'Vide le lister
75      Call lvwfournisseur.ListItems.Clear
            
80      If m_bPieceInutile = True Or m_bChangementFRS = True Then
85        Call CalculerPrixReel(Trim$(lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE)))
90      Else
95        If m_bRecherchePiece = True Then
100         Call CalculerPrixReel(Trim$(lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM)))
105       Else
110         Call CalculerPrixReel(Trim$(lvwPieces.SelectedItem.SubItems(I_COL_PIECES_NO_ITEM)))
115       End If
120     End If
      
125     Call rstFRS.Open("SELECT IDFRS FROM GRB_Fournisseur WHERE NomFournisseur = 'FOURNI PAR LE CLIENT'", g_connData, adOpenDynamic, adLockOptimistic)
      
130     iNoClient = rstFRS.Fields("IDFRS")

135     Call rstFRS.Close
140     Set rstFRS = Nothing
      
145     If m_bPieceInutile = True Or m_bChangementFRS = True Then
150       Call rstPieceFRS.Open("SELECT GRB_PiecesFRS.*, GRB_Fournisseur.NomFournisseur FROM GRB_PiecesFRS INNER JOIN GRB_Fournisseur ON GRB_PiecesFRS.IDFRS = GRB_Fournisseur.IDFRS WHERE PIECE = '" & Trim$(Replace(lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE), "'", "''")) & "' AND Type = 'M' ORDER BY CDbl(PrixReel)", g_connData, adOpenDynamic, adLockOptimistic)
155     Else
160       If m_bRecherchePiece = True Then
165         Call rstPieceFRS.Open("SELECT GRB_PiecesFRS.*, GRB_Fournisseur.NomFournisseur FROM GRB_PiecesFRS INNER JOIN GRB_Fournisseur ON GRB_PiecesFRS.IDFRS = GRB_Fournisseur.IDFRS WHERE PIECE = '" & Trim$(Replace(lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM), "'", "''")) & "' AND Type = 'M' ORDER BY CDbl(PrixReel)", g_connData, adOpenDynamic, adLockOptimistic)
170       Else
175         Call rstPieceFRS.Open("SELECT GRB_PiecesFRS.*, GRB_Fournisseur.NomFournisseur FROM GRB_PiecesFRS INNER JOIN GRB_Fournisseur ON GRB_PiecesFRS.IDFRS = GRB_Fournisseur.IDFRS WHERE PIECE = '" & Trim$(Replace(lvwPieces.SelectedItem.SubItems(I_COL_PIECES_NO_ITEM), "'", "''")) & "' AND Type = 'M' ORDER BY CDbl(PrixReel)", g_connData, adOpenDynamic, adLockOptimistic)
180       End If
185     End If
  
        'Tant il y a des fournisseur de la piece, ajoute dans lister
190     Do While Not rstPieceFRS.EOF
195       If m_bPieceInutile = True Then
200         If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_BRUN Then
205           If rstPieceFRS.Fields("IDFRS") = iNoClient Then
210             Call rstPieceFRS.MoveNext

215             If rstPieceFRS.EOF Then
220               Exit Do
225             End If
230           End If
235         End If
240       End If
          
          'on change la couleur de l'enregistrement selon la devise monétaire.
          'CAN = rouge, USA ou ESP = bleu
245       If rstPieceFRS.Fields("DeviseMonétaire") = "CAN" Then
250         sDevise = "CAN"
255         lColor = COLOR_NOIR
260       Else
265         If rstPieceFRS.Fields("DeviseMonétaire") = "USA" Then
270           sDevise = "USA"
275           lColor = COLOR_BLEU
280         Else
285           sDevise = "SPA"
290           lColor = COLOR_BLEU
295         End If
300       End If
       
305       Set itmFRS = lvwfournisseur.ListItems.Add
       
310       itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = " "
315       itmFRS.SubItems(I_COL_FRS_ESCOMPTE) = " "
320       itmFRS.SubItems(I_COL_FRS_PRIX_NET) = " "
325       itmFRS.SubItems(I_COL_FRS_PRIX_SP) = " "
          
          'Nom du FRS
330       itmFRS.Text = rstPieceFRS.Fields("NomFournisseur")
              
335       itmFRS.Tag = rstPieceFRS.Fields("IDFRS")

340       itmFRS.ForeColor = lColor
        
          'Personne ressource
345       If Trim(rstPieceFRS.Fields("PERS_RESS")) <> vbNullString Then
350         Call rstContact.Open("SELECT NomContact FROM GRB_Contact WHERE IDContact = " & rstPieceFRS.Fields("PERS_RESS"), g_connData, adOpenDynamic, adLockOptimistic)
            
355         If Not rstContact.EOF Then
360           itmFRS.SubItems(I_COL_FRS_PERS_RESS) = rstContact.Fields("NomContact")

365           itmFRS.ListSubItems(I_COL_FRS_PERS_RESS).ForeColor = lColor
370         End If

375         Call rstContact.Close
380       End If
                     
          'Date
385       If Not IsNull(rstPieceFRS.Fields("Date")) Then
390         itmFRS.SubItems(I_COL_FRS_DATE) = rstPieceFRS.Fields("Date")
395       Else
400         itmFRS.SubItems(I_COL_FRS_DATE) = vbNullString
405       End If

410       itmFRS.ListSubItems(I_COL_FRS_DATE).ForeColor = lColor
                          
          'Entrer par
415       If Not IsNull(rstPieceFRS.Fields("Entrer_par")) Then
420         itmFRS.SubItems(I_COL_FRS_ENTRER_PAR) = rstPieceFRS.Fields("Entrer_Par")
425       Else
430         itmFRS.SubItems(I_COL_FRS_ENTRER_PAR) = vbNullString
435       End If

440       itmFRS.ListSubItems(I_COL_FRS_ENTRER_PAR).ForeColor = lColor

          'Valide
445       If Not IsNull(rstPieceFRS.Fields("Valide")) Then
450         itmFRS.SubItems(I_COL_FRS_VALIDE) = rstPieceFRS.Fields("Valide")
455       Else
460         itmFRS.SubItems(I_COL_FRS_VALIDE) = vbNullString
465       End If

470       itmFRS.ListSubItems(I_COL_FRS_VALIDE).ForeColor = lColor
                             
          'Prix listé
475       If rstPieceFRS.Fields("PRIX_LIST") <> vbNullString Then
480         itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = Conversion(Round(Replace(rstPieceFRS.Fields("PRIX_LIST"), ".", ","), 4), MODE_ARGENT, 4)
485       End If
      
490       itmFRS.ListSubItems(I_COL_FRS_PRIX_LIST).ForeColor = lColor
      
          'Escompte
495       If rstPieceFRS.Fields("ESCOMPTE") <> vbNullString Then
500         itmFRS.SubItems(I_COL_FRS_ESCOMPTE) = Conversion(Replace(rstPieceFRS.Fields("ESCOMPTE"), "_", vbNullString) * 100, MODE_POURCENT)
505       End If
           
510       itmFRS.ListSubItems(I_COL_FRS_ESCOMPTE).ForeColor = lColor

          'Prix net
515       If rstPieceFRS.Fields("PRIX_NET") <> vbNullString Then
520         itmFRS.SubItems(I_COL_FRS_PRIX_NET) = Conversion(Round(Replace(rstPieceFRS.Fields("PRIX_NET"), ".", ","), 4), MODE_ARGENT, 4)
525       End If
            
530       itmFRS.ListSubItems(I_COL_FRS_PRIX_NET).ForeColor = lColor
            
          'Prix spécial
535       If rstPieceFRS.Fields("PRIX_SP") <> vbNullString Then
540         itmFRS.SubItems(I_COL_FRS_PRIX_SP) = Conversion(Round(rstPieceFRS.Fields("PRIX_SP"), 4), MODE_ARGENT, 4)
545       End If
       
550       itmFRS.ListSubItems(I_COL_FRS_PRIX_SP).ForeColor = lColor
       
          'Quoter
555       If rstPieceFRS.Fields("QUOTER") = True Then
560         itmFRS.SubItems(I_COL_FRS_QUOTER) = "Oui"
565       Else
570         itmFRS.SubItems(I_COL_FRS_QUOTER) = "Non"
575       End If
     
580       itmFRS.ListSubItems(I_COL_FRS_QUOTER).ForeColor = lColor
     
585       If rstPieceFRS.Fields("IDFRS") = 717 Then 'Si le fournisseur est "SOLUTION GRB Inc."
590         Set rstInv = New ADODB.Recordset

595         Call rstInv.Open("SELECT * FROM GRB_InventaireMec WHERE TRIM(NoItem) = '" & Trim$(Replace(rstPieceFRS.Fields("PIECE"), "'", "''")) & "'", g_connData, adOpenForwardOnly, adLockReadOnly)
        
600         If Not rstInv.EOF Then
605           If Not IsNull(rstInv.Fields("QuantitéStock")) Then
610             itmFRS.SubItems(I_COL_FRS_STOCK) = rstInv.Fields("QuantitéStock")
615           Else
620             itmFRS.SubItems(I_COL_FRS_STOCK) = 0
625           End If
630         End If

635         Call rstInv.Close
640         Set rstInv = Nothing
645       End If

          'Pour garder en mémoire le prix d'origine, je le mets dans le
          'tag de la colonne Prix Listé
650       If itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = vbNullString Then
655         itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = " "
660       End If
      
665       If rstPieceFRS.Fields("PRIX_NET") <> vbNullString Then
670         If rstPieceFRS.Fields("PRIX_LIST") = "0,00" Or rstPieceFRS.Fields("PRIX_LIST") = "0" Then
675           itmFRS.ListSubItems(I_COL_FRS_PRIX_LIST).Tag = Replace(rstPieceFRS.Fields("PRIX_NET"), ".", ",")
680         Else
685           itmFRS.ListSubItems(I_COL_FRS_PRIX_LIST).Tag = Replace(rstPieceFRS.Fields("PRIX_LIST"), ".", ",")
690         End If
695       Else
700         itmFRS.ListSubItems(I_COL_FRS_PRIX_LIST).Tag = Replace(rstPieceFRS.Fields("PRIX_SP"), ".", ",")
705       End If

710       If itmFRS.SubItems(I_COL_FRS_ENTRER_PAR) = vbNullString Then
715         itmFRS.SubItems(I_COL_FRS_ENTRER_PAR) = " "
720       End If

725       itmFRS.ListSubItems(I_COL_FRS_ENTRER_PAR).Tag = rstPieceFRS.Fields("NoEnreg")

730       If itmFRS.SubItems(I_COL_FRS_PERS_RESS) = "" Then
735         itmFRS.SubItems(I_COL_FRS_PERS_RESS) = " "
740       End If
      
745       itmFRS.ListSubItems(I_COL_FRS_PERS_RESS).Tag = sDevise

750       Call rstPieceFRS.MoveNext
755     Loop
    
        'Ferme la table
760     Call rstPieceFRS.Close
765     Set rstPieceFRS = Nothing

770     Set rstContact = Nothing

        'Modification temporaire
775     If m_bPieceInutile = False Then
780       If lvwSoumission.ListItems.count > 0 Then
785         If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor <> COLOR_BRUN Then
790           bAjouterDP = True
795         End If
800       Else
805         bAjouterDP = True
810       End If
815     Else
820       If m_bChangementFRS = True Then
825         If lvwSoumission.ListItems.count > 0 Then
830           If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor <> COLOR_BRUN Then
835             bAjouterDP = True
840           End If
845         Else
850           bAjouterDP = True
855         End If
860       End If
865     End If

870     If bAjouterDP = True Then
875       Set itmFRS = lvwfournisseur.ListItems.Add

880       itmFRS.Text = "CHOISIR ULTÉRIEUREMENT"
885       itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = " "
890       itmFRS.SubItems(I_COL_FRS_ESCOMPTE) = " "
895       itmFRS.SubItems(I_COL_FRS_PRIX_NET) = " "
900       itmFRS.SubItems(I_COL_FRS_PRIX_SP) = " "
905     End If
        
910     Exit Sub

AfficherErreur:

915     woups "frmProjSoumMec", "RemplirListViewFournisseur", Err, Erl
End Sub

Private Sub RemplirListViewPieces()

5       On Error GoTo AfficherErreur

        'Rempli le listview des pièces selon la catégorie de pièce choisit
10      Dim rstPieces  As ADODB.Recordset
15      Dim itmPieces  As ListItem
20      Dim sCategorie As String
25      Dim sTri       As String
30      Dim sOrderBy   As String
35      Dim bDebut     As Boolean
40      Dim iIndex     As Integer
  
45      sTri = m_sTri
  
50      Select Case cmbTri.ListIndex
          Case I_CMB_PIECE_GRB: sOrderBy = "PIECE_GRB"
55        Case I_CMB_PIECE:     sOrderBy = "PIECE"
60        Case I_CMB_FABRICANT: sOrderBy = "FABRICANT"
65        Case I_CMB_DESCR_FR:  sOrderBy = "DESC_FR"
70        Case I_CMB_DESCR_EN:  sOrderBy = "DESC_EN"
75      End Select
  
        'Il faut vider le ListView avant de le remplir
80      Call lvwPieces.ListItems.Clear
  
85      sCategorie = Replace(cmbPieces.Text, "'", "''")
  
        'On ouvre un recordset selon la table choisie
90      Set rstPieces = New ADODB.Recordset
        
95      Call rstPieces.Open("SELECT * FROM GRB_CatalogueMec WHERE CATEGORIE = '" & sCategorie & "' ORDER BY " & sOrderBy, g_connData, adOpenDynamic, adLockOptimistic)
    
100     iIndex = 1
    
        'Tant que ce n'est pas la fin des enregistrements
105     Do While Not rstPieces.EOF
110       If rstPieces.Fields("PIECE") <> vbNullString And rstPieces.Fields("FABRICANT") <> vbNullString Then
            'Si il y a une recherche à faire
115         If sTri <> vbNullString Then
120           bDebut = False
      
              'Selon la colonne
125           Select Case m_iCol
                'Si c'est la colonne PIECE_GRB
                Case I_COL_PIECES_PIECE_GRB:
                  'Si la PIECE_GRB contient la recherche
130               If InStr(1, UCase(rstPieces.Fields("PIECE_GRB")), UCase(sTri)) > 0 Then
                    'On met la variable à true pour l'ajouter au début
135                 bDebut = True
140               End If
                      
                'Si c'est la colonne No. d'item
145             Case I_COL_PIECES_NO_ITEM:
                  'Si le no. d'item contient la recherche
150               If InStr(1, UCase(rstPieces.Fields("PIECE")), UCase(sTri)) > 0 Then
                    'On met la variable à true pour l'ajouter au début
155                 bDebut = True
160               End If
        
                'Si c'est la colonne Manufacturier
165             Case I_COL_PIECES_MANUFACT:
                  'Si le manufacturier contient la recherche
170               If InStr(1, UCase(rstPieces.Fields("FABRICANT")), UCase(sTri)) > 0 Then
                    'On met la variable à true pour l'ajouter au début
175                 bDebut = True
180               End If
            
                'Si c'est la colonne No. d'item
185             Case I_COL_PIECES_DESCR_FR:
                  'Si la description française contient la recherche
190               If InStr(1, UCase(rstPieces.Fields("DESC_FR")), UCase(sTri)) > 0 Then
                    'On met la variable à true pour l'ajouter au début
195                 bDebut = True
200               End If
            
                'Si c'est la colonne No. d'item
205             Case I_COL_PIECES_DESCR_EN:
                  'Si la description anglaise contient la recherche
210               If InStr(1, UCase(rstPieces.Fields("DESC_EN")), UCase(sTri)) > 0 Then
                    'On met la variable à true pour l'ajouter au début
215                 bDebut = True
220               End If
225           End Select
        
230           If bDebut = True Then
235             Set itmPieces = lvwPieces.ListItems.Add(iIndex)
          
240             iIndex = iIndex + 1
245           Else
250             Set itmPieces = lvwPieces.ListItems.Add
255           End If
260         Else
265           Set itmPieces = lvwPieces.ListItems.Add
270         End If
                         
            'PIECE_GRB
275         If Not IsNull(rstPieces.Fields("PIECE_GRB")) Then
280           itmPieces.Text = rstPieces.Fields("PIECE_GRB")
285         Else
290           itmPieces.Text = vbNullString
295         End If
      
            'PIECE
300         If Not IsNull(rstPieces.Fields("PIECE")) Then
305           itmPieces.SubItems(I_COL_PIECES_NO_ITEM) = rstPieces.Fields("PIECE")
310         Else
315           itmPieces.SubItems(I_COL_PIECES_NO_ITEM) = vbNullString
320         End If
      
            'FABRICANT
325         If Not IsNull(rstPieces.Fields("FABRICANT")) Then
330           itmPieces.SubItems(I_COL_PIECES_MANUFACT) = rstPieces.Fields("FABRICANT")
335         Else
340           itmPieces.SubItems(I_COL_PIECES_MANUFACT) = vbNullString
345         End If
                 
            'DESCR_FR
350         If Not IsNull(rstPieces.Fields("DESC_FR")) Then
355           itmPieces.SubItems(I_COL_PIECES_DESCR_FR) = rstPieces.Fields("DESC_FR")
360         Else
365           itmPieces.SubItems(I_COL_PIECES_DESCR_FR) = vbNullString
370         End If
      
            'DESCR_EN
375         If Not IsNull(rstPieces.Fields("DESC_EN")) Then
380           itmPieces.SubItems(I_COL_PIECES_DESCR_EN) = rstPieces.Fields("DESC_EN")
385         Else
390           itmPieces.SubItems(I_COL_PIECES_DESCR_EN) = vbNullString
395         End If
400       End If
     
405       Call rstPieces.MoveNext
410     Loop
    
415     Call rstPieces.Close
420     Set rstPieces = Nothing

425     Exit Sub

AfficherErreur:

430     woups "frmProjSoumMec", "RemplirListViewPieces", Err, Erl
End Sub

Private Function TrouverIndexSection(ByVal sSousSection As String) As Integer

5       On Error GoTo AfficherErreur

        'recherche la section et l'ajouter si elle n'a pas été trouvée
10      Dim iCompteur         As Integer
15      Dim iIndex            As Integer
20      Dim iTagSection       As Integer
25      Dim iIDSection        As Integer
30      Dim iIndexSect        As Integer
35      Dim iIndexSSection    As Integer
40      Dim bTrouverSect      As Boolean
45      Dim bTrouverSSect     As Boolean
50      Dim bTrouverIndexItem As Boolean
55      Dim sTagSousSection   As String
60      Dim itmSoum           As ListItem
    
        'Si la variable sSousSection = PAS DE SOUS-SECTION
65      If sSousSection = S_PAS_SOUS_SECTION Then
          'On l'initialise à rien
70        sSousSection = vbNullString
          'On met le tag à PAS DE SOUS-SECTION
75        sTagSousSection = S_PAS_SOUS_SECTION
80      Else
85        sTagSousSection = sSousSection
90      End If
    
        'Si le listview n'est pas vide
95      If lvwSoumission.ListItems.count > 0 Then
          'Pour chaque élément du listview
100        For iCompteur = 1 To lvwSoumission.ListItems.count
            'Si c'est écrit le nom de la section dans la colonne Piece
105         If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) = cmbSections.Text Then
              'La section a été trouvée
110           bTrouverSect = True
        
              'On stock l'index de la section
115           iIndexSect = iCompteur
        
              'On commence à rechercher la sous-section à l'index suivant
120           iCompteur = iCompteur + 1
        
              'Tant que le tag du listItem est égal à l'index de la section
125           Do While lvwSoumission.ListItems(iCompteur).Tag = cmbSections.ItemData(cmbSections.ListIndex)
                'Si c'est écrit le nom de la sous-section dans la colonne Description
130             If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_DESCR) = sSousSection And lvwSoumission.ListItems(iCompteur).ListSubItems(I_COL_SOUM_DESCR).Bold = True Then
                  'La sous-section a été trouvée
135               bTrouverSSect = True
            
                  'On stock l'index du premier enregistrement de la section
140               iIndex = iCompteur + 1
            
145               Exit Do
150             End If
          
155             iCompteur = iCompteur + 1
          
                'Si le compteur est plus grand que le listItems.Count, il ne faut pas repasser
                'dans la boucle
160             If iCompteur > lvwSoumission.ListItems.count Then
165               Exit Do
170             End If
175           Loop
          
180           Exit For
185         End If
190       Next
195     Else
200       bTrouverSect = False
205     End If
    
210     If bTrouverSect = False Then
          'Ajoute la section
    
          'Si il y a des enregistrements dans le listview
215       If lvwSoumission.ListItems.count > 0 Then
            'Pour chaque élément du listview
220         For iCompteur = 1 To lvwSoumission.ListItems.count
              'Si ce n'est pas une section
225           If lvwSoumission.ListItems(iCompteur).Tag <> vbNullString Then
                'iTagSection est égal à l'ordre de la section du listitem
230             iTagSection = lvwSoumission.ListItems(iCompteur).ListSubItems(I_COL_SOUM_MANUFACT).Tag
                'iIDSection est égal à l'ordre de la section du combo
235             iIDSection = cmbSections.ListIndex + 1
          
                'Le premier enregistrement est 2 puisque 1 c'est une section
240             If iCompteur = 2 Then
                  'Si l'index de la section du combo est plus petit que
                  'l'index de la section du ListItem
245               If iIDSection < iTagSection Then
250                 iIndex = 1
           
255                 Exit For
260               End If
265             Else
270               If iCompteur = lvwSoumission.ListItems.count Then
                    'Si l'index de la section du combo est plus grand que l'index
                    'de la section du ListItem
275                 If iIDSection > iTagSection Then
280                   iIndex = iCompteur + 1
              
285                   Exit For
290                 End If
295               Else
300                 If lvwSoumission.ListItems(iCompteur + 1).Tag <> vbNullString Then
                      'Si l'index de la section du combo est plus grand que l'index
                      'de la section du ListItem et que l'index de la section du combo
                      'est plus petit que l'index de la section du ListItem suivant
305                   If iIDSection > iTagSection And iIDSection < lvwSoumission.ListItems(iCompteur + 1).ListSubItems(I_COL_SOUM_MANUFACT).Tag Then
310                     iIndex = iCompteur + 1
         
315                     Exit For
320                   End If
325                 Else
                      'Si l'index de la section du combo est plus grand que l'index
                      'de la section du ListItem et que l'index de la section du combo
                      'est plus petit que l'index de 2 ListItem plus loin
330                   If iIDSection > iTagSection And iIDSection < lvwSoumission.ListItems(iCompteur + 2).ListSubItems(I_COL_SOUM_MANUFACT).Tag Then
335                     iIndex = iCompteur + 1
340                   End If
345                 End If
350               End If
355             End If
360           End If
365         Next
370       Else
375         iIndex = 1
380       End If
      
          'On ajoute la section au bon index
385       Set itmSoum = lvwSoumission.ListItems.Add(iIndex)
     
390       itmSoum.SubItems(I_COL_SOUM_PIECE) = cmbSections.Text
395       itmSoum.ListSubItems(I_COL_SOUM_PIECE).Bold = True
       
400       itmSoum.SubItems(I_COL_SOUM_MANUFACT) = " "
405       itmSoum.ListSubItems(I_COL_SOUM_MANUFACT).Tag = cmbSections.ListIndex + 1
    
410       iIndex = iIndex + 1
    
415       Call ValeurParDefaut(itmSoum)
    
          'On l'ajoute la sous-section à l'index suivant
420       iIndexSSection = AjouterSousSection(iIndex, sTagSousSection)
          
425       iIndex = iIndexSSection
430     Else
          'Si la sous-section n'a pas été trouvé dans le listview
435       If bTrouverSSect = False Then
            'On l'ajoute à l'index suivant la section
440         iIndexSSection = AjouterSousSection(iIndexSect + 1, sTagSousSection)
        
445         iIndex = iIndexSSection
450       End If
455     End If
   
        'Pour trouver le dernier élément de la sous-section
  
        'Pour chaque élément du listview à partir de l'index du premier élément de la sous-section
460     For iCompteur = iIndex To lvwSoumission.ListItems.count
          'Si on trouve une autre sous-section ou une autre section
465       If lvwSoumission.ListItems(iCompteur).Tag <> cmbSections.ItemData(cmbSections.ListIndex) Or lvwSoumission.ListItems(iCompteur).ListSubItems(I_COL_SOUM_PIECE).Tag <> sTagSousSection Then
470         bTrouverIndexItem = True
            
            'On l'ajoute à cette index
475         iIndex = iCompteur
      
480         Exit For
485       End If
490     Next
  
        'Si la fin de la sous-section n'a pas été, il faut l'ajouter à la fin
495     If bTrouverIndexItem = False Then
500       iIndex = lvwSoumission.ListItems.count + 1
505     End If
  
510     TrouverIndexSection = iIndex

515     Exit Function

AfficherErreur:

520     woups "frmProjSoumMec", "TrouverIndexSection", Err, Erl
End Function

Private Function AjouterSousSection(ByVal iIndexSection As Integer, ByVal sSousSection As String) As Integer

5       On Error GoTo AfficherErreur

        'Méthode qui sert à ajouter une sous-section
10      Dim itmSoum               As ListItem
15      Dim iCompteur             As Integer
20      Dim bTrouverIndexSSection As Boolean
25      Dim iIndex                As Integer
30      Dim sTag                  As String
      
35      If sSousSection = S_PAS_SOUS_SECTION Then
40        sSousSection = vbNullString
45        sTag = S_PAS_SOUS_SECTION
50      Else
55        sTag = sSousSection
60      End If
 
65      If sTag <> S_PAS_SOUS_SECTION Then
          'Pour chaque élément du listview
70        For iCompteur = iIndexSection To lvwSoumission.ListItems.count
            'Si le tag du ListItem est différent du IDSection de la section
75          If lvwSoumission.ListItems(iCompteur).Tag <> cmbSections.ItemData(cmbSections.ListIndex) Then
80            bTrouverIndexSSection = True
     
85            iIndex = iCompteur
      
90            Exit For
95          End If
100       Next
  
          'Si l'emplacement de la sous-section n'a pas été trouvée
105       If bTrouverIndexSSection = False Then
            'On la place à la fin
110         iIndex = lvwSoumission.ListItems.count + 1
115       End If
120     Else
125       iIndex = iIndexSection
130     End If
  
135     Set itmSoum = lvwSoumission.ListItems.Add(iIndex)
  
140     Call ValeurParDefaut(itmSoum)
      
        'On met le nom de la sous-section dans la colonne Description
145     itmSoum.SubItems(I_COL_SOUM_DESCR) = sSousSection
150     itmSoum.ListSubItems(I_COL_SOUM_DESCR).Bold = True
  
        'On met le nom de la sous-section dans le tag de la colonne Piece
'155     itmSoum.ListSubItems(I_COL_SOUM_PIECE).Tag = sTag
      
        'On met l'ID de la section dans le tag du listitem
155     itmSoum.Tag = cmbSections.ItemData(cmbSections.ListIndex)
  
        'On ne peut pas écrire dans le tag si vide
160     itmSoum.SubItems(I_COL_SOUM_MANUFACT) = " "
  
        'Ordre de la section
165     itmSoum.ListSubItems(I_COL_SOUM_MANUFACT).Tag = cmbSections.ListIndex + 1
  
170     AjouterSousSection = iIndex + 1

175     Exit Function

AfficherErreur:

180     woups "frmProjSoumMec", "AjouterSousSection", Err, Erl
End Function

Private Sub AjouterNegatifDansListView(ByVal dblQuantite As Double, ByVal sSousSection As String)

5       On Error GoTo AfficherErreur

10      Dim rstProjet   As ADODB.Recordset
15      Dim iIndex      As Integer
20      Dim iCompteur   As Integer
25      Dim iIDSection  As Integer
30      Dim iTagSection As Integer
35      Dim iIndexSel   As Integer
40      Dim bSelected   As Boolean
45      Dim dblTotalQte As Double
50      Dim lColor      As Long
55      Dim bQteOK      As Boolean
60      Dim sNoProjet   As String
65      Dim sPrixList   As String
70      Dim sEscompte   As String
75      Dim sPrixNet    As String
80      Dim itmSoum     As ListItem

85      Set rstProjet = New ADODB.Recordset
  
90      If Right$(txtNoProjSoum.Text, 2) >= 60 And Right$(txtNoProjSoum.Text, 2) <= 98 Then
95        sNoProjet = Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & m_sLiaison

100       If m_bRecherchePiece = True Then
105         Call rstProjet.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoProjet & "' AND NumItem = '" & lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM) & "' AND IDFRS = " & lvwfournisseur.SelectedItem.Tag, g_connData, adOpenDynamic, adLockOptimistic)
110       Else
115         Call rstProjet.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoProjet & "' AND NumItem = '" & lvwPieces.SelectedItem.SubItems(I_COL_PIECES_NO_ITEM) & "' AND IDFRS = " & lvwfournisseur.SelectedItem.Tag, g_connData, adOpenDynamic, adLockOptimistic)
120       End If
125     End If

130     If Not rstProjet.EOF Then
135       Do While Not rstProjet.EOF
140         dblTotalQte = dblTotalQte + rstProjet.Fields("Qté")

145         Call rstProjet.MoveNext
150       Loop

155       If dblTotalQte >= Abs(dblQuantite) Then
160         bQteOK = True
165       End If
170     Else
175       Call MsgBox("La pièce n'existe pas dans le projet " & sNoProjet, vbOKOnly, "Erreur")

180       Call rstProjet.Close
185       Set rstProjet = Nothing

190       Exit Sub
195     End If
  
200     If bQteOK = True Then
205       Call rstProjet.MovePrevious

210       sPrixList = rstProjet.Fields("Prix_List")
215       sEscompte = rstProjet.Fields("Escompte")
220       sPrixNet = rstProjet.Fields("Prix_Net")
225     Else
230       If m_bRecherchePiece = True Then
235         Call MsgBox("Il n'y a pas assez de " & lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM) & " dans le projet " & sNoProjet & " pour en enlever " & Abs(dblQuantite) & "!", vbOKOnly, "Erreur")
240       Else
245         Call MsgBox("Il n'y a pas assez de " & lvwPieces.SelectedItem.SubItems(I_COL_PIECES_NO_ITEM) & " dans le projet " & sNoProjet & " pour en enlever " & Abs(dblQuantite) & "!", vbOKOnly, "Erreur")
250       End If

255       Call rstProjet.Close
260       Set rstProjet = Nothing

265       Exit Sub
270     End If

275     Call rstProjet.Close
280     Set rstProjet = Nothing

285     bSelected = False
  
        'S'il y a des items dans le ListView
290     If lvwSoumission.ListItems.count > 0 Then
          'Si ce n'est pas le premier qui est sélectionné
          '(le premier est sélectionné par défaut)
295       If lvwSoumission.SelectedItem.Index > 1 Then
300         bSelected = True

305         iIndexSel = lvwSoumission.SelectedItem.Index
310       End If
315     End If

        'si le premier est sélectionné, on le considère comme si aucun n'est sélectionné
320     If bSelected = False Then
325       iIndex = TrouverIndexSection(sSousSection)
330     Else
          'Sinon, on l'ajoute à l'endroit sélectionné
335       iIndex = iIndexSel
340     End If

345     Set itmSoum = lvwSoumission.ListItems.Add(iIndex)

350     itmSoum.Checked = True

        'Quantité
355     itmSoum.Text = dblQuantite

360     If lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_QUOTER) = "Oui" Then
365       itmSoum.Text = itmSoum.Text & "*"
370       itmSoum.ForeColor = COLOR_VERT
375       itmSoum.Bold = True
380     Else
385       itmSoum.ForeColor = COLOR_NOIR
390       itmSoum.Bold = False
395     End If

        'On met l'id de la section dans le tag du listItem
400     itmSoum.Tag = cmbSections.ItemData(cmbSections.ListIndex) 'IDSection
                                                                                                         
        'No d'item
405     If m_bRecherchePiece = True Then
410       itmSoum.SubItems(I_COL_SOUM_PIECE) = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM)
415     Else
420       itmSoum.SubItems(I_COL_SOUM_PIECE) = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_NO_ITEM)
425     End If

430     itmSoum.ListSubItems(I_COL_SOUM_PIECE).ForeColor = lColor

        'On met le nom de la sous-section dans le tag du no d'item
435     itmSoum.ListSubItems(I_COL_SOUM_PIECE).Tag = sSousSection
  
        'On met la description en francais dans la colonne et la description en anglais
        'dans le tag
440     If m_eLangage = ANGLAIS Then
445       If m_bRecherchePiece = True Then
450         itmSoum.SubItems(I_COL_SOUM_DESCR) = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_DESCR_EN)
455         itmSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_DESCR_FR)
460       Else
465         itmSoum.SubItems(I_COL_SOUM_DESCR) = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_DESCR_EN)
470         itmSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_DESCR_FR)
475       End If
480     Else
485       If m_bRecherchePiece = True Then
490         itmSoum.SubItems(I_COL_SOUM_DESCR) = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_DESCR_FR)
495         itmSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_DESCR_EN)
500       Else
505         itmSoum.SubItems(I_COL_SOUM_DESCR) = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_DESCR_FR)
510         itmSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_DESCR_EN)
515       End If
520     End If

525     itmSoum.ListSubItems(I_COL_SOUM_DESCR).ForeColor = lColor
          
        'On met le nom du fabricant dans la colonne et l'ordre de la section dans le tag
530     If m_bRecherchePiece = True Then
535       itmSoum.SubItems(I_COL_SOUM_MANUFACT) = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_MANUFACT)
540     Else
545       itmSoum.SubItems(I_COL_SOUM_MANUFACT) = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_MANUFACT)
550     End If

555     itmSoum.ListSubItems(I_COL_SOUM_MANUFACT).Tag = cmbSections.ListIndex + 1 'Ordre de la section
  
560     itmSoum.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = lColor

        'Prix listé
565     If Trim$(sPrixList) = vbNullString Then
570       itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion("0", MODE_ARGENT, 4)
575     Else
580       itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(sPrixList, MODE_ARGENT, 4)
585       itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = sPrixList
590     End If
      
595     itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = lColor
       
        'Si il y a un prix net, on le met l'escompte et le prix net sinon, on prend le prix
        'spécial pour mettre dans le prix net
600     If Trim$(sEscompte) <> vbNullString Then
605       itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = sEscompte
610     Else
615       itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion("0", MODE_POURCENT)
620     End If
      
625     itmSoum.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = lColor

630     If Trim$(sPrixNet) <> vbNullString Then
635       itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(sPrixNet, MODE_ARGENT, 4)
640     Else
645       itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion("0", MODE_ARGENT, 4)
650     End If

655     itmSoum.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = lColor

660     itmSoum.SubItems(I_COL_SOUM_DISTRIB) = lvwfournisseur.SelectedItem.Text
665     itmSoum.ListSubItems(I_COL_SOUM_DISTRIB).Tag = lvwfournisseur.SelectedItem.Tag
    
670     itmSoum.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = lColor
                
        'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
675     itmSoum.SubItems(I_COL_SOUM_TOTAL) = Conversion(Round(Replace(itmSoum.Text, "*", vbNullString) * itmSoum.SubItems(I_COL_SOUM_PRIX_NET) * CSng(m_sProfit), 2), MODE_ARGENT, 2)
     
680     itmSoum.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = lColor
      
        'Pour le profit, c'est le prix total - (prix net * quantité)
685     itmSoum.SubItems(I_COL_SOUM_PROFIT) = Conversion(CStr(Round(itmSoum.SubItems(I_COL_SOUM_TOTAL) - (itmSoum.SubItems(I_COL_SOUM_PRIX_NET) * Replace(itmSoum.Text, "*", vbNullString)), 2)), MODE_ARGENT)

690     itmSoum.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = lColor

695     If itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = vbNullString Then
700       itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = " "
705     End If

710     Exit Sub

AfficherErreur:

715     woups "frmProjSoumMec", "AjouterDansListViewSoumission", Err, Erl
End Sub

Private Sub AjouterDansListViewSoumission(ByVal dblQuantite As Double, ByVal sSousSection As String)

5       On Error GoTo AfficherErreur

10      Dim rstConfig   As ADODB.Recordset
15      Dim itmSoum     As ListItem
20      Dim iIndex      As Integer
25      Dim iCompteur   As Integer
30      Dim iIDSection  As Integer
35      Dim iTagSection As Integer
40      Dim iIndexSel   As Integer
45      Dim bSelected   As Boolean
50      Dim bAjouter    As Boolean
55      Dim sDistrib    As String
60      Dim sTauxUSA    As String
65      Dim sTauxSPA    As String
70      Dim lColor      As Long

75      Set rstConfig = New ADODB.Recordset

80      Call rstConfig.Open("SELECT TauxAmericain, TauxEspagnol FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)
  
85      sTauxUSA = rstConfig.Fields("TauxAmericain")
90      sTauxSPA = rstConfig.Fields("TauxEspagnol")

95      Call rstConfig.Close
100     Set rstConfig = Nothing

105     bSelected = False
  
        'Si il y a des items dans le ListView
110     If lvwSoumission.ListItems.count > 0 Then
          'Si ce n'est pas le premier qui est sélectionné
          '(le premier est sélectionné par défaut)
115       If lvwSoumission.SelectedItem.Index > 1 Then
120         bSelected = True
      
125         iIndexSel = lvwSoumission.SelectedItem.Index
130       End If
135     End If
 
        'si le premier est sélectionné, on le considère comme si aucun n'est sélectionné
140     If bSelected = False Then
145       iIndex = TrouverIndexSection(sSousSection)
150     Else
          'Sinon, on l'ajoute à l'endroit sélectionné
155       iIndex = iIndexSel
160     End If

165     If VerifierSiPieceExiste(sSousSection, iIndex) = True Then
170       If sSousSection <> "PAS DE SOUS-SECTION" Then
175         If MsgBox("Cette pièce existe déjà dans la sous-section '" & sSousSection & "' de la section '" & cmbSections.Text & "'." & vbNewLine & _
                      "Voulez-vous l'ajouter à nouveau ?", vbYesNo) = vbYes Then
180           bAjouter = True
185         Else
190           bAjouter = False
195         End If
200       Else
205         If MsgBox("Cette pièce existe déjà dans la sous-section vide de la section '" & cmbSections.Text & "'." & vbNewLine & _
                      "Voulez-vous l'ajouter à nouveau ?", vbYesNo) = vbYes Then
210           bAjouter = True
215         Else
220           bAjouter = False
225         End If
230       End If
235     Else
240       bAjouter = True
245     End If
                    
250     If bAjouter = True Then
255       Set itmSoum = lvwSoumission.ListItems.Add(iIndex)
  
260       itmSoum.Checked = True
    
          'Quantité
265       itmSoum.Text = dblQuantite
    
270       If lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_QUOTER) = "Oui" Then
275         itmSoum.Text = itmSoum.Text & "*"
280         itmSoum.ForeColor = COLOR_VERT
285         itmSoum.Bold = True
290       Else
295         itmSoum.ForeColor = COLOR_NOIR
300         itmSoum.Bold = False
305       End If
    
310       If lvwfournisseur.SelectedItem.Text = "CHOISIR ULTÉRIEUREMENT" Then
315         lColor = COLOR_MAGENTA
320       Else
325         lColor = COLOR_NOIR
330       End If
   
          'On met l'id de la section dans le tag du listItem
335       itmSoum.Tag = cmbSections.ItemData(cmbSections.ListIndex) 'IDSection
                                                                                                           
          'No d'item
340       If m_bRecherchePiece = True Then
345         itmSoum.SubItems(I_COL_SOUM_PIECE) = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM)
350       Else
355         itmSoum.SubItems(I_COL_SOUM_PIECE) = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_NO_ITEM)
360       End If

365       itmSoum.ListSubItems(I_COL_SOUM_PIECE).ForeColor = lColor
    
          'On met le nom de la sous-section dans le tag du no d'item
370       itmSoum.ListSubItems(I_COL_SOUM_PIECE).Tag = sSousSection
    
          'On met la description en francais dans la colonne et la description en anglais
          'dans le tag
375       If m_eLangage = ANGLAIS Then
380         If m_bRecherchePiece = True Then
385           itmSoum.SubItems(I_COL_SOUM_DESCR) = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_DESCR_EN)
390           itmSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_DESCR_FR)
395         Else
400           itmSoum.SubItems(I_COL_SOUM_DESCR) = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_DESCR_EN)
405           itmSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_DESCR_FR)
410         End If
415       Else
420         If m_bRecherchePiece = True Then
425           itmSoum.SubItems(I_COL_SOUM_DESCR) = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_DESCR_FR)
430           itmSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_DESCR_EN)
435         Else
440           itmSoum.SubItems(I_COL_SOUM_DESCR) = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_DESCR_FR)
445           itmSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_DESCR_EN)
450         End If
455       End If

460       itmSoum.ListSubItems(I_COL_SOUM_DESCR).ForeColor = lColor
            
          'On met le nom du fabricant dans la colonne et l'ordre de la section dans le tag
465       If m_bRecherchePiece = True Then
470         itmSoum.SubItems(I_COL_SOUM_MANUFACT) = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_MANUFACT)
475       Else
480          itmSoum.SubItems(I_COL_SOUM_MANUFACT) = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_MANUFACT)
485       End If
  
490       itmSoum.ListSubItems(I_COL_SOUM_MANUFACT).Tag = cmbSections.ListIndex + 1 'Ordre de la section

495       itmSoum.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = lColor
    
          'Prix listé
500       If Trim$(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_LIST)) = vbNullString Then
505         itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion("0", MODE_ARGENT, 4)
510       Else
515         If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "USA" Then
520           itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_LIST)) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
525         Else
530           If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "SPA" Then
535             itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_LIST)) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
540           Else
545             itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_LIST), MODE_ARGENT, 4)
550           End If
555         End If
560       End If
       
565       itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PRIX_LIST).Tag
          
570       itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = lColor
         
          'Si il y a un prix net, on le met l'escompte et le prix net sinon, on prend le prix
          'spécial pour mettre dans le prix net
575       If Trim$(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_NET)) <> vbNullString Then
580         If Trim$(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_ESCOMPTE)) <> "" Then
585           itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_ESCOMPTE), MODE_POURCENT)
590         Else
595           itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion(0, MODE_POURCENT)
600         End If
  
605         If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "USA" Then
610           itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_NET)) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
615         Else
620           If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "SPA" Then
625             itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_NET)) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
630           Else
635             itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_NET), MODE_ARGENT, 4)
640           End If
645         End If
650       Else
655         If Trim$(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_SP)) <> vbNullString Then
660           If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "USA" Then
665             itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_SP)) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
670           Else
675             If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "USA" Then
680               itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_SP)) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
685             Else
690               itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_SP), MODE_ARGENT, 4)
695             End If
700           End If
705         Else
710           itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion("0", MODE_POURCENT)
715           itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion("0", MODE_ARGENT, 4)
720         End If
725       End If
       
730       itmSoum.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = lColor
735       itmSoum.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = lColor
     
          'On met le fournisseur dans la colonne et l'id dans le tag
740       If lvwfournisseur.SelectedItem.Text = "CHOISIR ULTÉRIEUREMENT" Then
745         sDistrib = vbNullString
750       Else
755         sDistrib = lvwfournisseur.SelectedItem.Text
760       End If
  
765       itmSoum.SubItems(I_COL_SOUM_DISTRIB) = sDistrib
770       itmSoum.ListSubItems(I_COL_SOUM_DISTRIB).Tag = lvwfournisseur.SelectedItem.Tag

775       itmSoum.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = lColor
            
          'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
780       itmSoum.SubItems(I_COL_SOUM_TOTAL) = Conversion(CStr(Round(Replace(itmSoum.Text, "*", vbNullString) * itmSoum.SubItems(I_COL_SOUM_PRIX_NET) * CSng(m_sProfit), 2)), MODE_ARGENT)
  
785       itmSoum.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = lColor
        
          'Pour le profit, c'est le prix total - (prix net * quantité)
790       itmSoum.SubItems(I_COL_SOUM_PROFIT) = Conversion(CStr(Round(itmSoum.SubItems(I_COL_SOUM_TOTAL) - (itmSoum.SubItems(I_COL_SOUM_PRIX_NET) * Replace(itmSoum.Text, "*", vbNullString)), 2)), MODE_ARGENT)
  
795       itmSoum.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = lColor
  
800       If m_eType = TYPE_PROJET Then
805         itmSoum.SubItems(I_COL_SOUM_FACTURATION) = " "
810       End If
  
815       Call itmSoum.EnsureVisible
820     End If

825     Exit Sub

AfficherErreur:

830     woups "frmProjSoumMec", "AjouterDansListViewSoumission", Err, Erl
End Sub

Private Function VerifierSiPieceExiste(ByVal sSousSection As String, ByVal iIndex As Integer) As Boolean

5       On Error GoTo AfficherErreur

10      Dim bExiste   As Boolean
15      Dim iCompteur As Integer
20      Dim sPiece    As String

25      If lvwSoumission.ListItems.count >= iIndex Then
30        If lvwSoumission.ListItems(iIndex).ListSubItems(I_COL_SOUM_PIECE).Tag = sSousSection Then
35          iCompteur = iIndex
40        Else
45          iCompteur = iIndex - 1
50        End If
55      Else
60        iCompteur = iIndex - 1
65      End If

70      Do While lvwSoumission.ListItems(iCompteur).ListSubItems(I_COL_SOUM_PIECE).Tag = sSousSection
75        iCompteur = iCompteur - 1
80      Loop

85      iCompteur = iCompteur + 1

90      If iCompteur <= lvwSoumission.ListItems.count Then
95        If m_bRecherchePiece = True Then
100         sPiece = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM)
105       Else
110         sPiece = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_NO_ITEM)
115       End If

120       Do While lvwSoumission.ListItems(iCompteur).ListSubItems(I_COL_SOUM_PIECE).Tag = sSousSection
125         If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) = sPiece Then
130           bExiste = True

135           Exit Do
140         End If

145         iCompteur = iCompteur + 1

150         If iCompteur > lvwSoumission.ListItems.count Then
155           Exit Do
160         End If
165       Loop
170     Else
175       bExiste = False
180     End If

185     VerifierSiPieceExiste = bExiste

190     Exit Function

AfficherErreur:

195     woups "frmProjSoumMec", "VerifierSiPieceExiste", Err, Erl
End Function

Private Function VerifierEmplacement(ByVal iIndexSelection As Integer) As Boolean

5       On Error GoTo AfficherErreur

        'Vérifie si l'emplacement pour ajouter une pièce est valide
10      Dim itmSoum As ListItem
  
15      Set itmSoum = lvwSoumission.ListItems(iIndexSelection)
  
20      If itmSoum.Tag = vbNullString Then
25        Set itmSoum = lvwSoumission.ListItems(iIndexSelection - 1)
30      End If
  
        'Si la section est correcte
35      If itmSoum.Tag = cmbSections.ItemData(cmbSections.ListIndex) Then
40        VerifierEmplacement = True
45      Else
50        VerifierEmplacement = False
55      End If

60      Exit Function

AfficherErreur:

65      woups "frmProjSoumMec", "VerifierEmplacement", Err, Erl
End Function

Private Sub ValeurParDefaut(ByVal itmSoumission As ListItem)

5       On Error GoTo AfficherErreur

        'Méthode pour mettre une valeur par défaut dans quelques colonnes de lvwSoumission.
        'Si ces colonnes sont vides, elles restent blanches lors de la sélection
10      If m_bDroitPrix = True Then
15        itmSoumission.SubItems(I_COL_SOUM_PRIX_LIST) = " "
20        itmSoumission.SubItems(I_COL_SOUM_ESCOMPTE) = " "
25        itmSoumission.SubItems(I_COL_SOUM_PRIX_NET) = " "
30        itmSoumission.SubItems(I_COL_SOUM_TOTAL) = " "
35        itmSoumission.SubItems(I_COL_SOUM_PROFIT) = " "
40      End If

45      Exit Sub

AfficherErreur:

50      woups "frmProjSoumMec", "ValeurParDefaut", Err, Erl
End Sub

Private Sub RemplirListViewProjSoum(ByVal sNoProjSoum As String)

5       On Error GoTo AfficherErreur

        'Remplis les pièces de la soumission avec la BD
10      Dim rstProjSoum   As ADODB.Recordset
15      Dim rstSection    As ADODB.Recordset
20      Dim rstFRS        As ADODB.Recordset
25      Dim itmProjSoum   As ListItem
30      Dim bPremierEnr   As Boolean
35      Dim bBold         As Boolean
40      Dim iOrdreSection As Integer
45      Dim sSousSection  As String
50      Dim sSection      As String
55      Dim lColor        As Long
  
60      Call lvwSoumission.ListItems.Clear
  
65      bPremierEnr = True
  
70      Set rstProjSoum = New ADODB.Recordset
  
75      If m_eType = TYPE_PROJET Then
80        Call rstProjSoum.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoProjSoum & "' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
85      Else
90        Call rstProjSoum.Open("SELECT * FROM GRB_Soumission_Pieces WHERE IDSoumission = '" & sNoProjSoum & "' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
95      End If
    
100     If m_eLangage = ANGLAIS Then
105       sSection = "NomSectionEN"
110     Else
115       sSection = "NomSectionFR"
120     End If

125     Set rstSection = New ADODB.Recordset
130     Set rstFRS = New ADODB.Recordset

135     Do While Not rstProjSoum.EOF
140       Set itmProjSoum = lvwSoumission.ListItems.Add
            
          'Si c'est le premier enregistrement, il faut ajouter la section et la sous-section
145       If bPremierEnr = True Then
150         iOrdreSection = rstProjSoum.Fields("OrdreSection")
155         sSousSection = rstProjSoum.Fields("SousSection")
              
            'Pour avoir le nom de la section
160         Call rstSection.Open("SELECT " & sSection & " FROM GRB_SoumProjSectionMec WHERE IDSection = " & rstProjSoum.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
        
            'Ajout du nom de la section
165         If Not IsNull(rstSection.Fields(sSection)) Then
170           itmProjSoum.SubItems(I_COL_SOUM_PIECE) = rstSection.Fields(sSection)
175         Else
180           itmProjSoum.SubItems(I_COL_SOUM_PIECE) = vbNullString
185         End If
        
190         itmProjSoum.ListSubItems(I_COL_SOUM_PIECE).Bold = True
       
195         Call ValeurParDefaut(itmProjSoum)
                      
200         Call rstSection.Close
          
205         Set itmProjSoum = lvwSoumission.ListItems.Add
        
            'Ajout du nom de la sous-section
210         If sSousSection = S_PAS_SOUS_SECTION Then
215           itmProjSoum.SubItems(I_COL_SOUM_DESCR) = vbNullString
220         Else
225           itmProjSoum.SubItems(I_COL_SOUM_DESCR) = sSousSection
230         End If
        
            'Le tag ne peut pas être remplis si la colonne est vide
235         itmProjSoum.SubItems(I_COL_SOUM_MANUFACT) = " "
      
240         itmProjSoum.ListSubItems(I_COL_SOUM_MANUFACT).Tag = iOrdreSection
               
245         itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Bold = True
        
250         itmProjSoum.Tag = rstProjSoum.Fields("IDSection")
        
255         Call ValeurParDefaut(itmProjSoum)
        
260         Set itmProjSoum = lvwSoumission.ListItems.Add
        
265         bPremierEnr = False
270       Else
            'Si c'est pas le premier enregistrement, il faut vérifier avec l'ancienne section
275         If iOrdreSection <> rstProjSoum.Fields("OrdreSection") Then
280           iOrdreSection = rstProjSoum.Fields("OrdreSection")
        
285           Call rstSection.Open("SELECT " & sSection & " FROM GRB_SoumProjSectionMec WHERE IDSection = " & rstProjSoum.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
          
290           If Not IsNull(rstSection.Fields(sSection)) Then
295             itmProjSoum.SubItems(I_COL_SOUM_PIECE) = rstSection.Fields(sSection)
300           Else
305             itmProjSoum.SubItems(I_COL_SOUM_PIECE) = vbNullString
310           End If
          
315           itmProjSoum.ListSubItems(I_COL_SOUM_PIECE).Bold = True
          
320           Call ValeurParDefaut(itmProjSoum)
          
325           Call rstSection.Close
                
330           Set itmProjSoum = lvwSoumission.ListItems.Add
          
335           sSousSection = rstProjSoum.Fields("SousSection")
          
340           If sSousSection = S_PAS_SOUS_SECTION Then
345             itmProjSoum.SubItems(I_COL_SOUM_DESCR) = vbNullString
350           Else
355             itmProjSoum.SubItems(I_COL_SOUM_DESCR) = rstProjSoum.Fields("SousSection")
360           End If
          
              'Le tag ne peut pas être remplis si la colonne est vide
365           itmProjSoum.SubItems(I_COL_SOUM_MANUFACT) = " "
        
370           itmProjSoum.ListSubItems(I_COL_SOUM_MANUFACT).Tag = iOrdreSection
          
375           itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Bold = True
          
380           Call ValeurParDefaut(itmProjSoum)
          
385           itmProjSoum.Tag = rstProjSoum.Fields("IDSection")
          
390           Set itmProjSoum = lvwSoumission.ListItems.Add
395         Else
              'il faut vérifier avec l'ancienne sous-section
400           If sSousSection <> rstProjSoum.Fields("SousSection") Then
405             sSousSection = rstProjSoum.Fields("SousSection")
            
410             If sSousSection = S_PAS_SOUS_SECTION Then
415               itmProjSoum.SubItems(I_COL_SOUM_DESCR) = vbNullString
420             Else
425               itmProjSoum.SubItems(I_COL_SOUM_DESCR) = sSousSection
430             End If
                            
435             itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Bold = True
         
440             Call ValeurParDefaut(itmProjSoum)
          
445             itmProjSoum.Tag = rstProjSoum.Fields("IDSection")
            
                'Le tag ne peut pas être remplis si la colonne est vide
450             itmProjSoum.SubItems(I_COL_SOUM_MANUFACT) = " "
        
455             itmProjSoum.ListSubItems(I_COL_SOUM_MANUFACT).Tag = iOrdreSection
          
460             Set itmProjSoum = lvwSoumission.ListItems.Add
465           End If
470         End If
475       End If
  
480       If rstProjSoum.Fields("PieceExtraNonChargeable") = True Then
485         lColor = COLOR_ROSE
490         bBold = True
495       Else
500         If rstProjSoum.Fields("PieceExtraChargeable") = True Then
505           lColor = COLOR_BLEU
510           bBold = True
515         Else
520           If rstProjSoum.Fields("CommandeAnnulée") = True Then
525             lColor = COLOR_VERT_FORET
530             bBold = True
535           Else
540             If rstProjSoum.Fields("Retour") = True Then
545               lColor = COLOR_ROUGE
550               bBold = False
555             Else
560               If rstProjSoum.Fields("Commandé") = True Then
565                 lColor = COLOR_ORANGE     'COLOR_ORANGE
570                 bBold = False
575               Else
580                 If rstProjSoum.Fields("Recu") = True Then
585                   lColor = COLOR_GRIS 'Gris
590                   bBold = False
595                 Else
600                   If rstProjSoum.Fields("IDFRS") = 0 And rstProjSoum.Fields("NumItem") <> "Texte" And rstProjSoum.Fields("NumItem") <> "Text" Then
605                     lColor = COLOR_MAGENTA
610                     bBold = False
615                   Else
620                     If rstProjSoum.Fields("MatérielInutile") = True Then
625                       lColor = COLOR_BRUN
630                       bBold = False
635                     Else
640                       lColor = COLOR_NOIR
645                       bBold = False
650                     End If
655                   End If
660                 End If
665               End If
670             End If
675           End If
680         End If
685       End If

          'On met l'ID de la section dans le tag
690       itmProjSoum.Tag = rstProjSoum.Fields("IDSection")
      
695       If rstProjSoum.Fields("Visible") = True Then
700         itmProjSoum.Checked = True
705       Else
710         itmProjSoum.Checked = False
715       End If
      
          'Quantité
720       If Not IsNull(rstProjSoum.Fields("Qté")) Then
725         itmProjSoum.Text = rstProjSoum.Fields("Qté")
730       Else
735         itmProjSoum.Text = vbNullString
740       End If
    
745       If rstProjSoum.Fields("Quoté") = True Then
750         itmProjSoum.Text = itmProjSoum.Text & "*"
755         itmProjSoum.ForeColor = COLOR_VERT
760         itmProjSoum.Bold = True
765       Else
770         itmProjSoum.ForeColor = COLOR_NOIR
775         itmProjSoum.Bold = False
780       End If

785       If m_eType = TYPE_PROJET Then
790         If g_bModificationProjetsMec = True Then
              'Facturation
795           If Not IsNull(rstProjSoum.Fields("Facturation")) Then
800             If Trim(rstProjSoum.Fields("Facturation")) <> "" Then
805               itmProjSoum.SubItems(I_COL_SOUM_FACTURATION) = rstProjSoum.Fields("Facturation")
810             Else
815               itmProjSoum.SubItems(I_COL_SOUM_FACTURATION) = " "
820             End If
825           Else
830             itmProjSoum.SubItems(I_COL_SOUM_FACTURATION) = " "
835           End If
840         End If
845       End If
      
          'Numéro d'item
850       If Not IsNull(rstProjSoum.Fields("NumItem")) Then
855         If rstProjSoum.Fields("NumItem") = "Texte" Or rstProjSoum.Fields("NumItem") = "Text" Then
860           If m_eLangage = ANGLAIS Then
865             itmProjSoum.SubItems(I_COL_SOUM_PIECE) = "Text"
870           Else
875             itmProjSoum.SubItems(I_COL_SOUM_PIECE) = "Texte"
880           End If
885         Else
890           itmProjSoum.SubItems(I_COL_SOUM_PIECE) = rstProjSoum.Fields("NumItem")
895         End If
900       Else
905         itmProjSoum.SubItems(I_COL_SOUM_PIECE) = vbNullString
910       End If
   
915       itmProjSoum.ListSubItems(I_COL_SOUM_PIECE).ForeColor = lColor

920       itmProjSoum.ListSubItems(I_COL_SOUM_PIECE).Bold = bBold
     
          'On met le nom de la sous-section dans le tag du numéro d'item
925       itmProjSoum.ListSubItems(I_COL_SOUM_PIECE).Tag = rstProjSoum.Fields("SousSection")
    
930       If itmProjSoum.SubItems(I_COL_SOUM_PIECE) = "Texte" Or itmProjSoum.SubItems(I_COL_SOUM_PIECE) = "Text" Then
935         itmProjSoum.SubItems(I_COL_SOUM_DESCR) = rstProjSoum.Fields("DESC_FR")
940       Else
945         If m_eLangage = ANGLAIS Then
              'Description en anglais
950           If Not IsNull(rstProjSoum.Fields("DESC_EN")) Then
955             itmProjSoum.SubItems(I_COL_SOUM_DESCR) = rstProjSoum.Fields("DESC_EN")
960           Else
965             itmProjSoum.SubItems(I_COL_SOUM_DESCR) = vbNullString
970           End If

              'On met la description en francais dans le tag de la description en anglais
975           If Not IsNull(rstProjSoum.Fields("DESC_FR")) Then
980             itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = rstProjSoum.Fields("DESC_FR")
985           Else
990             itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = vbNullString
995           End If
1000        Else
              'Description en francais
1005          If Not IsNull(rstProjSoum.Fields("DESC_FR")) Then
1010            itmProjSoum.SubItems(I_COL_SOUM_DESCR) = rstProjSoum.Fields("DESC_FR")
1015          Else
1020            itmProjSoum.SubItems(I_COL_SOUM_DESCR) = vbNullString
1025          End If

              'On met la description en anglais dans le tag de la description en francais
1030          If Not IsNull(rstProjSoum.Fields("Desc_EN")) Then
1035            itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = rstProjSoum.Fields("Desc_EN")
1040          Else
1045            itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = vbNullString
1050          End If
1055        End If
1060      End If
    
1065      itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).ForeColor = lColor
1070      itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Bold = bBold
      
          'Fabricant
1075      If Not IsNull(rstProjSoum.Fields("Manufact")) Then
1080        itmProjSoum.SubItems(I_COL_SOUM_MANUFACT) = rstProjSoum.Fields("Manufact")
1085      Else
1090        itmProjSoum.SubItems(I_COL_SOUM_MANUFACT) = vbNullString
1095      End If
    
1100      itmProjSoum.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = lColor
  
1105      itmProjSoum.ListSubItems(I_COL_SOUM_MANUFACT).Bold = bBold
      
          'On met l'ordre de la section dans le tag du fabricant
1110      itmProjSoum.ListSubItems(I_COL_SOUM_MANUFACT).Tag = iOrdreSection
     
          'Prix listé
1115      If m_bDroitPrix = True Then
1120        If Trim(rstProjSoum.Fields("Prix_List")) <> vbNullString Then
1125          itmProjSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(rstProjSoum.Fields("Prix_list"), MODE_ARGENT, 4)
1130        Else
1135          itmProjSoum.SubItems(I_COL_SOUM_PRIX_LIST) = " "
1140        End If
     
1145        itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = lColor

1150        itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).Bold = bBold
      
1155        itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = rstProjSoum.Fields("PrixOrigine")
       
            'Escompte
1160        If Trim(rstProjSoum.Fields("Escompte")) <> vbNullString Then
1165          itmProjSoum.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion(rstProjSoum.Fields("Escompte"), MODE_POURCENT)
1170        Else
1175          itmProjSoum.SubItems(I_COL_SOUM_ESCOMPTE) = " "
1180        End If
      
1185        itmProjSoum.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = lColor

1190        itmProjSoum.ListSubItems(I_COL_SOUM_ESCOMPTE).Bold = bBold
     
            'Prix net
1195        If Trim(rstProjSoum.Fields("Prix_net")) <> vbNullString Then
1200          itmProjSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(rstProjSoum.Fields("Prix_net"), MODE_ARGENT, 4)
1205        Else
1210          itmProjSoum.SubItems(I_COL_SOUM_PRIX_NET) = " "
1215        End If
            
1220        itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = lColor

1225        itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_NET).Bold = bBold
            
1230        If m_eType = TYPE_PROJET Then
1235          If Not IsNull(rstProjSoum.Fields("DateRéception")) Then
1240            itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_NET).Tag = rstProjSoum.Fields("DateRéception")
1245          End If
1250        End If
           
            'Fournisseur
1255        If Not IsNull(rstProjSoum.Fields("IDFRS")) And rstProjSoum.Fields("IDFRS") <> "0" Then
1260          If itmProjSoum.SubItems(I_COL_SOUM_PIECE) <> "Texte" And itmProjSoum.SubItems(I_COL_SOUM_PIECE) <> "Text" Then
1265            Call rstFRS.Open("SELECT NomFournisseur FROM GRB_Fournisseur WHERE IDFRS = " & rstProjSoum.Fields("IDFRS"), g_connData, adOpenDynamic, adLockOptimistic)
    
                'On affiche le nom dans la colonne
1270            itmProjSoum.SubItems(I_COL_SOUM_DISTRIB) = rstFRS.Fields("NomFournisseur")
         
1275            itmProjSoum.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = lColor

1280            itmProjSoum.ListSubItems(I_COL_SOUM_DISTRIB).Bold = bBold
        
                'On affiche l'Id dans le tag
1285            itmProjSoum.ListSubItems(I_COL_SOUM_DISTRIB).Tag = rstProjSoum.Fields("IDFRS")
        
1290            Call rstFRS.Close
1295          End If
1300        Else
1305          itmProjSoum.SubItems(I_COL_SOUM_DISTRIB) = vbNullString
1310          itmProjSoum.ListSubItems(I_COL_SOUM_DISTRIB).Tag = 0
1315        End If
      
            'Prix total
1320        If Trim(rstProjSoum.Fields("Prix_total")) <> vbNullString Then
1325          itmProjSoum.SubItems(I_COL_SOUM_TOTAL) = Conversion(Round(rstProjSoum.Fields("Prix_total"), 2), MODE_ARGENT)
1330        Else
1335          itmProjSoum.SubItems(I_COL_SOUM_TOTAL) = " "
1340        End If
    
1345        itmProjSoum.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = lColor

1350        itmProjSoum.ListSubItems(I_COL_SOUM_TOTAL).Bold = bBold

            'Profit
1355        If Trim(rstProjSoum.Fields("Profit_argent")) <> vbNullString Then
1360          itmProjSoum.SubItems(I_COL_SOUM_PROFIT) = Conversion(Round(rstProjSoum.Fields("Profit_Argent"), 2), MODE_ARGENT)
1365        Else
1370          itmProjSoum.SubItems(I_COL_SOUM_PROFIT) = " "
1375        End If

1380        If m_eType = TYPE_PROJET Then
1385          If rstProjSoum.Fields("PieceExtraNonChargeable") = True Then
1390            itmProjSoum.ListSubItems(I_COL_SOUM_PROFIT).Tag = "EXTRA"
1395          End If
1400        End If
      
1405        itmProjSoum.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = lColor

1410        itmProjSoum.ListSubItems(I_COL_SOUM_PROFIT).Bold = bBold

1415        If Not IsNull(rstProjSoum.Fields("Commentaire")) Then
1420          itmProjSoum.SubItems(I_COL_SOUM_COMMENTAIRE) = rstProjSoum.Fields("Commentaire")
1425        Else
1430          itmProjSoum.SubItems(I_COL_SOUM_COMMENTAIRE) = vbNullString
1435        End If

1440        itmProjSoum.ListSubItems(I_COL_SOUM_COMMENTAIRE).ForeColor = lColor

1445        itmProjSoum.ListSubItems(I_COL_SOUM_COMMENTAIRE).Bold = bBold

1450        If m_eType = TYPE_PROJET Then
1455          If Not IsNull(rstProjSoum.Fields("DateCommande")) Then
1460            If Trim(rstProjSoum.Fields("DateCommande")) <> "" Then
1465              itmProjSoum.SubItems(I_COL_SOUM_DATE_COMMANDE) = rstProjSoum.Fields("DateCommande")
1470            Else
1475              itmProjSoum.SubItems(I_COL_SOUM_DATE_COMMANDE) = " "
1480            End If
1485          Else
1490            itmProjSoum.SubItems(I_COL_SOUM_DATE_COMMANDE) = " "
1495          End If

1500          itmProjSoum.ListSubItems(I_COL_SOUM_DATE_COMMANDE).ForeColor = lColor

1505          itmProjSoum.ListSubItems(I_COL_SOUM_DATE_COMMANDE).Bold = bBold

1510          itmProjSoum.ListSubItems(I_COL_SOUM_DATE_COMMANDE).Tag = rstProjSoum.Fields("NoRetour")

1515          If Not IsNull(rstProjSoum.Fields("DateRequise")) Then
1520            If Trim(rstProjSoum.Fields("DateRequise")) <> "" Then
1525              itmProjSoum.SubItems(I_COL_SOUM_DATE_REQUISE) = rstProjSoum.Fields("DateRequise")
1530            Else
1535              itmProjSoum.SubItems(I_COL_SOUM_DATE_REQUISE) = " "
1540            End If
1545          Else
1550            itmProjSoum.SubItems(I_COL_SOUM_DATE_REQUISE) = " "
1555          End If

1560          itmProjSoum.ListSubItems(I_COL_SOUM_DATE_REQUISE).ForeColor = lColor

1565          itmProjSoum.ListSubItems(I_COL_SOUM_DATE_REQUISE).Bold = bBold

1570          itmProjSoum.ListSubItems(I_COL_SOUM_DATE_REQUISE).Tag = rstProjSoum.Fields("DateRetour")

1575          If Not IsNull(rstProjSoum.Fields("NomCommande")) Then
1580            itmProjSoum.SubItems(I_COL_SOUM_NOM_COMMANDE) = rstProjSoum.Fields("NomCommande")
1585          Else
1590            itmProjSoum.SubItems(I_COL_SOUM_NOM_COMMANDE) = vbNullString
1595          End If

1600          itmProjSoum.ListSubItems(I_COL_SOUM_NOM_COMMANDE).ForeColor = lColor

1605          itmProjSoum.ListSubItems(I_COL_SOUM_NOM_COMMANDE).Bold = bBold

1610          If Not IsNull(rstProjSoum.Fields("NoSéquentiel")) Then
1615            itmProjSoum.SubItems(I_COL_SOUM_NO_SEQUENTIEL) = rstProjSoum.Fields("NoSéquentiel")
1620          Else
1625            itmProjSoum.SubItems(I_COL_SOUM_NO_SEQUENTIEL) = vbNullString
1630          End If

1635          itmProjSoum.ListSubItems(I_COL_SOUM_NO_SEQUENTIEL).ForeColor = lColor

1640          itmProjSoum.ListSubItems(I_COL_SOUM_NO_SEQUENTIEL).Bold = bBold

1645          If Not IsNull(rstProjSoum.Fields("Provenance")) Then
1650            If Trim(rstProjSoum.Fields("Provenance")) <> "" Then
1655              itmProjSoum.SubItems(I_COL_SOUM_PROVENANCE) = Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 3) & "-" & rstProjSoum.Fields("Provenance")
1660            Else
1665              itmProjSoum.SubItems(I_COL_SOUM_PROVENANCE) = ""
1670            End If
1675          Else
1680            itmProjSoum.SubItems(I_COL_SOUM_PROVENANCE) = ""
1685          End If

1690          itmProjSoum.ListSubItems(I_COL_SOUM_PROVENANCE).ForeColor = lColor
  
1695          itmProjSoum.ListSubItems(I_COL_SOUM_PROVENANCE).Bold = bBold
1700        Else
1705          If Not IsNull(rstProjSoum.Fields("Provenance")) Then
1710            If Trim(rstProjSoum.Fields("Provenance")) <> "" Then
1715              itmProjSoum.SubItems(I_COL_SOUMISSION_PROV) = Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 3) & "-" & rstProjSoum.Fields("Provenance")
1720            Else
1725              itmProjSoum.SubItems(I_COL_SOUMISSION_PROV) = ""
1730            End If
1735          Else
1740            itmProjSoum.SubItems(I_COL_SOUMISSION_PROV) = ""
1745          End If

1750          itmProjSoum.ListSubItems(I_COL_SOUMISSION_PROV).ForeColor = lColor
  
1755          itmProjSoum.ListSubItems(I_COL_SOUMISSION_PROV).Bold = bBold
1760        End If
1765      Else
            'Fournisseur
1770        If Not IsNull(rstProjSoum.Fields("IDFRS")) And rstProjSoum.Fields("IDFRS") <> "0" Then
1775          If itmProjSoum.SubItems(I_COL_SOUM_SP_PIECE) <> "Texte" And itmProjSoum.SubItems(I_COL_SOUM_SP_PIECE) <> "Text" Then
1780            Call rstFRS.Open("SELECT NomFournisseur FROM GRB_Fournisseur WHERE IDFRS = " & rstProjSoum.Fields("IDFRS"), g_connData, adOpenDynamic, adLockOptimistic)
    
                'On affiche le nom dans la colonne
1785            itmProjSoum.SubItems(I_COL_SOUM_SP_DISTRIB) = rstFRS.Fields("NomFournisseur")
        
1790            itmProjSoum.ListSubItems(I_COL_SOUM_SP_DISTRIB).ForeColor = lColor

1795            itmProjSoum.ListSubItems(I_COL_SOUM_SP_DISTRIB).Bold = bBold
        
                'On affiche l'ID dans le tag
1800            itmProjSoum.ListSubItems(I_COL_SOUM_SP_DISTRIB).Tag = rstProjSoum.Fields("IDFRS")
        
1805            Call rstFRS.Close
1810          End If
1815        Else
1820          itmProjSoum.SubItems(I_COL_SOUM_SP_DISTRIB) = vbNullString
1825        End If

1830        If Not IsNull(rstProjSoum.Fields("Commentaire")) Then
1835          itmProjSoum.SubItems(I_COL_SOUM_SP_COMMENTAIRE) = rstProjSoum.Fields("Commentaire")
1840        Else
1845          itmProjSoum.SubItems(I_COL_SOUM_SP_COMMENTAIRE) = vbNullString
1850        End If
   
1855        itmProjSoum.ListSubItems(I_COL_SOUM_SP_COMMENTAIRE).ForeColor = lColor

1860        itmProjSoum.ListSubItems(I_COL_SOUM_SP_COMMENTAIRE).Bold = bBold

1865        If m_eType = TYPE_PROJET Then
1870          If Not IsNull(rstProjSoum.Fields("DateCommande")) Then
1875            If Trim(rstProjSoum.Fields("DateCommande")) <> "" Then
1880              itmProjSoum.SubItems(I_COL_SOUM_SP_DATE_COMMANDE) = rstProjSoum.Fields("DateCommande")
1885            Else
1890              itmProjSoum.SubItems(I_COL_SOUM_SP_DATE_COMMANDE) = " "
1895            End If
1900          Else
1905            itmProjSoum.SubItems(I_COL_SOUM_SP_DATE_COMMANDE) = " "
1910          End If

1915          itmProjSoum.ListSubItems(I_COL_SOUM_SP_DATE_COMMANDE).ForeColor = lColor

1920          itmProjSoum.ListSubItems(I_COL_SOUM_SP_DATE_COMMANDE).Bold = bBold

1925          itmProjSoum.ListSubItems(I_COL_SOUM_SP_DATE_COMMANDE).Tag = rstProjSoum.Fields("NoRetour")

1930          If Not IsNull(rstProjSoum.Fields("DateRequise")) Then
1935            If Trim(rstProjSoum.Fields("DateRequise")) <> "" Then
1940              itmProjSoum.SubItems(I_COL_SOUM_SP_DATE_REQUISE) = rstProjSoum.Fields("DateRequise")
1945            Else
1950              itmProjSoum.SubItems(I_COL_SOUM_SP_DATE_REQUISE) = " "
1955            End If
1960          Else
1965            itmProjSoum.SubItems(I_COL_SOUM_SP_DATE_REQUISE) = " "
1970          End If

1975          itmProjSoum.ListSubItems(I_COL_SOUM_SP_DATE_REQUISE).ForeColor = lColor

1980          itmProjSoum.ListSubItems(I_COL_SOUM_SP_DATE_REQUISE).Bold = bBold

1985          itmProjSoum.ListSubItems(I_COL_SOUM_SP_DATE_REQUISE).Tag = rstProjSoum.Fields("DateRetour")

1990          itmProjSoum.SubItems(I_COL_SOUM_SP_NOM_COMMANDE) = rstProjSoum.Fields("NomCommande")

1995          itmProjSoum.ListSubItems(I_COL_SOUM_SP_NOM_COMMANDE).ForeColor = lColor

2000          itmProjSoum.ListSubItems(I_COL_SOUM_SP_NOM_COMMANDE).Bold = bBold

2005          itmProjSoum.SubItems(I_COL_SOUM_SP_NO_SEQUENTIEL) = rstProjSoum.Fields("NoSéquentiel")

2010          itmProjSoum.ListSubItems(I_COL_SOUM_SP_NO_SEQUENTIEL).ForeColor = lColor

2015          itmProjSoum.ListSubItems(I_COL_SOUM_SP_NO_SEQUENTIEL).Bold = bBold

2025          If Not IsNull(rstProjSoum.Fields("Provenance")) Then
2030            If Trim(rstProjSoum.Fields("Provenance")) <> "" Then
2035              itmProjSoum.SubItems(I_COL_SOUM_SP_PROVENANCE) = Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 3) & "-" & rstProjSoum.Fields("Provenance")
2040            Else
2045              itmProjSoum.SubItems(I_COL_SOUM_SP_PROVENANCE) = ""
2050            End If
2055          Else
2060            itmProjSoum.SubItems(I_COL_SOUM_SP_PROVENANCE) = ""
2065          End If

2070          itmProjSoum.ListSubItems(I_COL_SOUM_SP_PROVENANCE).ForeColor = lColor

2075          itmProjSoum.ListSubItems(I_COL_SOUM_SP_PROVENANCE).Bold = bBold
2080        Else
2085          If Not IsNull(rstProjSoum.Fields("Provenance")) Then
2090            If Trim(rstProjSoum.Fields("Provenance")) <> "" Then
2095              itmProjSoum.SubItems(I_COL_SOUMISSION_SP_PROV) = Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 3) & "-" & rstProjSoum.Fields("Provenance")
2100            Else
2105              itmProjSoum.SubItems(I_COL_SOUMISSION_SP_PROV) = ""
2110            End If
2115          Else
2120            itmProjSoum.SubItems(I_COL_SOUMISSION_SP_PROV) = ""
2125          End If

2130          itmProjSoum.ListSubItems(I_COL_SOUMISSION_SP_PROV).ForeColor = lColor

2135          itmProjSoum.ListSubItems(I_COL_SOUMISSION_SP_PROV).Bold = bBold
2140        End If
2145      End If
      
2150      Call rstProjSoum.MoveNext
2155    Loop
   
2160    If lvwSoumission.ListItems.count > 0 Then
2165      Call Deselect

2170      lvwSoumission.ListItems(1).Selected = True
2175    End If
    
2180    Call rstProjSoum.Close
2185    Set rstProjSoum = Nothing

2190    Set rstFRS = Nothing
2195    Set rstSection = Nothing

2200    Call CalculerPrix

2205    Exit Sub

AfficherErreur:

2210    woups "frmProjSoumMec", "RemplirListViewProjSoum", Err, Erl, sNoProjSoum)
End Sub

Private Sub CalculerPrix()

5       On Error GoTo AfficherErreur

        'Méthode pour calculer le prix
10      Dim dblPrixPieces        As Double
15      Dim dblPrixTotal         As Double
20      Dim dblCommission        As Double
25      Dim dblTotalTemps        As Double
30      Dim dblProfit            As Double
35      Dim dblTotalManuel       As Double
40      Dim dblTotalImprevue     As Double
45      Dim dblGrandTotal        As Double
50      Dim dblTotalDessin       As Double
55      Dim dblTotalCoupe        As Double
60      Dim dblTotalMachinage    As Double
65      Dim dblTotalSoudure      As Double
70      Dim dblTotalAssemblage   As Double
75      Dim dblTotalPeinture     As Double
80      Dim dblTotalTest         As Double
85      Dim dblTotalInstallation As Double
90      Dim dblTotalFormation    As Double
95      Dim dblTotalGestion      As Double
100     Dim dblTotalShipping     As Double
105     Dim dblHebergement       As Double
110     Dim dblRepas             As Double
115     Dim dblTransport         As Double
120     Dim dblUniteMobile       As Double
125     Dim dblPrixEmballage     As Double
130     Dim dblTotalResteTemps   As Double
135     Dim bDemande             As Boolean
140     Dim iNbrePersonne        As Integer
145     Dim iCompteur            As Integer
        
        'Si ce n'est pas en mode affichage
150     If m_bModeAffichage = False Then
          'Pour chaque élément du listview
155       For iCompteur = 1 To lvwSoumission.ListItems.count
            'Si ce n'est pas une section
160         If lvwSoumission.ListItems(iCompteur).Tag <> vbNullString Then
              'Si ce n'est pas une sous-section
165           If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> vbNullString And lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> "Texte" And lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> "Text" Then
170             If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_DISTRIB) <> vbNullString Then
                  'On additionne le prix total
                  
175               If IsNumeric(lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_TOTAL)) And IsNumeric(lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PROFIT)) Then
180                 dblPrixPieces = dblPrixPieces + lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_TOTAL) - lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PROFIT)
185               Else
190                 Call MsgBox("La pièce " & lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) & " a un prix non numérique!", vbOKOnly, "Erreur")
195               End If
          
                  'On additionne le profit
200               If IsNumeric(Trim$(lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PROFIT))) = True Then
205                 dblProfit = dblProfit + lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PROFIT)
210               End If
215             Else
220               bDemande = True
225             End If
230           End If
235         End If
240       Next
    
          'Total des temps
245       dblTotalDessin = CDbl(m_sTempsDessin) * CDbl(m_sTauxDessin)
250       dblTotalCoupe = CDbl(m_sTempsCoupe) * CDbl(m_sTauxCoupe)
255       dblTotalMachinage = CDbl(m_sTempsMachinage) * CDbl(m_sTauxMachinage)
260       dblTotalSoudure = CDbl(m_sTempsSoudure) * CDbl(m_sTauxSoudure)
265       dblTotalAssemblage = CDbl(m_sTempsAssemblage) * CDbl(m_sTauxAssemblage)
270       dblTotalPeinture = CDbl(m_sTempsPeinture) * CDbl(m_sTauxPeinture)
275       dblTotalTest = CDbl(m_sTempsTest) * CDbl(m_sTauxTest)
280       dblTotalInstallation = CDbl(m_sTempsInstallation) * CDbl(m_sTauxInstallation)
285       dblTotalFormation = CDbl(m_sTempsFormation) * CDbl(m_sTauxFormation)
290       dblTotalGestion = CDbl(m_sTempsGestion) * CDbl(m_sTauxGestion)
295       dblTotalShipping = CDbl(m_sTempsShipping) * CDbl(m_sTauxShipping)
           
300       dblTotalTemps = dblTotalDessin + _
                          dblTotalCoupe + _
                          dblTotalMachinage + _
                          dblTotalSoudure + _
                          dblTotalAssemblage + _
                          dblTotalPeinture + _
                          dblTotalTest + _
                          dblTotalInstallation + _
                          dblTotalFormation + _
                          dblTotalGestion + _
                          dblTotalShipping
        
305       If m_eType = TYPE_PROJET Then
310         dblHebergement = 0
315         dblRepas = 0
320         dblTransport = 0
325         dblUniteMobile = 0
330       Else
335         iNbrePersonne = Int(m_sNbrePersonne)
           
340         Do While iNbrePersonne > 0
345           If iNbrePersonne >= 2 Then
350             dblHebergement = dblHebergement + CDbl(m_sTempsHebergement) * CDbl(m_sTauxHebergement2)
              
355             iNbrePersonne = iNbrePersonne - 2
360           Else
365             dblHebergement = dblHebergement + CDbl(m_sTempsHebergement) * CDbl(m_sTauxHebergement1)
             
370             iNbrePersonne = iNbrePersonne - 1
375           End If
380         Loop
      
385         dblRepas = CDbl(m_sTempsRepas) * CDbl(m_sTauxRepas) * CDbl(m_sNbrePersonne)
390         dblTransport = CDbl(m_sTempsTransport) * CDbl(m_sTauxTransport)
395         dblUniteMobile = CDbl(m_sTempsUniteMobile) * CDbl(m_sTauxUniteMobile)
400       End If

          'Correction d'un bug de Type Incompatible
405       If IsNumeric(m_sPrixEmballage) Then
410         dblPrixEmballage = CDbl(m_sPrixEmballage)
415       Else
420         dblPrixEmballage = 0
425       End If
      
430       dblTotalResteTemps = dblHebergement + dblRepas + dblTransport + dblUniteMobile + dblPrixEmballage
                                                              
435       If IsNumeric(txtPrixManuel.Text) Then
440         dblTotalManuel = CDbl(txtPrixManuel.Text)
445       Else
450         dblTotalManuel = 0
455       End If
                        
460       dblTotalImprevue = Round((dblPrixPieces + dblProfit) * CDbl(m_sImprevue), 2)
    
465       dblPrixTotal = dblPrixPieces + dblProfit + dblTotalTemps + dblTotalImprevue + dblTotalManuel + dblTotalResteTemps
                        
          'Le calcul de la commission est sur les manuels (Nbre de page * prix par pages * nbre de copies) + (prix des pièces * pourcentage d'imprévus) + total des temps
470       dblCommission = Round(dblPrixTotal * CDbl(m_sCommission), 2)
        
          'Le prix total est le calcul des manuels (Nbre de page * prix par pages * nbre de copies) + (prix des pièces * pourcentage d'imprévus) + total des temps + total de la commission
475       dblGrandTotal = dblPrixTotal + dblCommission
                
          'Format monétaires avec 2 chiffres après la virgule
480       txtTotalPieces.Text = Conversion(CStr(Round(dblPrixPieces, 2)), MODE_ARGENT)
485       txtTotalTemps.Text = Conversion(CStr(Round(dblTotalTemps, 2)), MODE_ARGENT)
490       txtPrixTotal.Text = Conversion(CStr(Round(dblGrandTotal, 2)), MODE_ARGENT)
          
495       If bDemande = True Then
500         txtPrixTotal.ForeColor = COLOR_JAUNE
505       Else
510         txtPrixTotal.ForeColor = COLOR_ROUGE
515       End If

520       txtImprevus.Text = Conversion(CStr(Round(dblTotalImprevue, 2)), MODE_ARGENT)
525       txtCommission.Text = Conversion(CStr(Round(dblCommission, 2)), MODE_ARGENT)
530       txtProfit.Text = Conversion(CStr(Round(dblProfit, 2)), MODE_ARGENT)
535     Else
540       For iCompteur = 1 To lvwSoumission.ListItems.count
            'Si ce n'est pas une section
545         If lvwSoumission.ListItems(iCompteur).Tag <> vbNullString Then
              'Si ce n'est pas une sous-section
550           If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> vbNullString And lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> "Texte" And lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> "Text" Then
555             If m_bDroitPrix = True Then
560               If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_DISTRIB) = vbNullString Then
565                 bDemande = True
                  
570                 Exit For
575               End If
580             End If
585           End If
590         End If
595       Next

600       If bDemande = True Then
605         txtPrixTotal.ForeColor = COLOR_JAUNE
610       Else
615         txtPrixTotal.ForeColor = COLOR_ROUGE
620       End If
625     End If

630     Exit Sub

AfficherErreur:

635     woups "frmProjSoumMec", "CalculerPrix", Err, Erl
End Sub

Private Sub CalculerTotalRecordset(ByVal sNoProjSoum As String)

5       On Error GoTo AfficherErreur

        'Méthode pour calculer le prix
10      Dim rstProjSoum          As ADODB.Recordset
15      Dim rstPiece             As ADODB.Recordset
20      Dim rstPunch             As ADODB.Recordset
25      Dim dblPrixPieces        As Double
30      Dim dblPrixTotal         As Double
35      Dim dblCommission        As Double
40      Dim dblTotalTemps        As Double
45      Dim dblProfit            As Double
50      Dim dblTotalManuel       As Double
55      Dim dblTotalImprevue     As Double
60      Dim dblGrandTotal        As Double
65      Dim dblTotalDessin       As Double
70      Dim dblTotalCoupe        As Double
75      Dim dblTotalMachinage    As Double
80      Dim dblTotalSoudure      As Double
85      Dim dblTotalAssemblage   As Double
90      Dim dblTotalPeinture     As Double
95      Dim dblTotalTest         As Double
100     Dim dblTotalInstallation As Double
105     Dim dblTotalFormation    As Double
110     Dim dblTotalGestion      As Double
115     Dim dblTotalShipping     As Double
120     Dim dblHebergement       As Double
125     Dim dblRepas             As Double
130     Dim dblTransport         As Double
135     Dim dblUniteMobile       As Double
140     Dim dblPrixEmballage     As Double
145     Dim dblTotalResteTemps   As Double
150     Dim sFilterNoProjet      As String
155     Dim sDateDebut           As String
160     Dim sDateFin             As String
165     Dim sTotal               As String

170     Set rstProjSoum = New ADODB.Recordset

175     If m_eType = TYPE_PROJET Then
180       Call rstProjSoum.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)
185     Else
190       Call rstProjSoum.Open("SELECT * FROM GRB_SoumissionMec WHERE IDSoumission = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)
195     End If

200     If Not rstProjSoum.EOF Then
205       Set rstPiece = New ADODB.Recordset

210       If m_eType = TYPE_PROJET Then
215         Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoProjSoum & "' AND Type = 'M'", g_connData, adOpenDynamic, adLockOptimistic)
220       Else
225         Call rstPiece.Open("SELECT * FROM GRB_Soumission_Pieces WHERE IDSoumission = '" & sNoProjSoum & "' AND Type = 'M'", g_connData, adOpenDynamic, adLockOptimistic)
230       End If

          'Pour chaque élément du recordset
235       Do While Not rstPiece.EOF
240         If Trim(rstPiece.Fields("Prix_total")) <> vbNullString Then
              'On additionne le prix total
245           dblPrixPieces = dblPrixPieces + rstPiece.Fields("Prix_total") - rstPiece.Fields("Profit_Argent")
          
              'On additionne le profit
250           dblProfit = dblProfit + rstPiece.Fields("Profit_Argent")
255         End If

260         Call rstPiece.MoveNext
265       Loop
    
          'Total des temps
270       If m_eType = TYPE_PROJET Then
275         sDateDebut = "TIMESERIAL(Left(GRB_Punch.HeureDébut,2),RIGHT(GRB_Punch.HeureDébut,2),0)"

280         sDateFin = "TIMESERIAL(Left(GRB_Punch.HeureFin,2),RIGHT(GRB_Punch.HeureFin,2),0)"

285         sTotal = "(SUM(" & sDateFin & " - " & sDateDebut & ")* 24) As Total"

290         If Right$(sNoProjSoum, 2) = "99" Then
295           sFilterNoProjet = "LEFT(NoProjet, 6) = '" & Left$(sNoProjSoum, 6) & "'"
300         Else
305           sFilterNoProjet = "NoProjet = '" & sNoProjSoum & "'"
310         End If

315         Set rstPunch = New ADODB.Recordset

320         Call rstPunch.Open("SELECT Type, " & sTotal & " FROM GRB_Punch WHERE " & sFilterNoProjet & " And HeureFin is not null AND HeureDébut is not null GROUP BY Type", g_connData, adOpenDynamic, adLockOptimistic)

325         dblTotalDessin = 0
330         dblTotalCoupe = 0
335         dblTotalMachinage = 0
340         dblTotalSoudure = 0
345         dblTotalAssemblage = 0
350         dblTotalPeinture = 0
355         dblTotalTest = 0
360         dblTotalInstallation = 0
365         dblTotalFormation = 0
370         dblTotalGestion = 0
375         dblTotalShipping = 0

380         Do While Not rstPunch.EOF
385           If Not IsNull(rstPunch.Fields("Total")) Then
390             Select Case rstPunch.Fields("Type")
                  Case "Dessin":
395                 If Not IsNull(rstProjSoum.Fields("TauxDessin")) Then
400                   dblTotalDessin = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxDessin"))
405                 Else
410                   dblTotalDessin = 0
415                 End If
                    
420               Case "Coupe":
425                 If Not IsNull(rstProjSoum.Fields("TauxCoupe")) Then
430                   dblTotalCoupe = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxCoupe"))
435                 Else
440                   dblTotalCoupe = 0
445                 End If
                    
450               Case "Machinage":
455                 If Not IsNull(rstProjSoum.Fields("TauxMachinage")) Then
460                   dblTotalMachinage = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxMachinage"))
465                 Else
470                   dblTotalMachinage = 0
475                 End If
                    
480               Case "Soudure":
485                 If Not IsNull(rstProjSoum.Fields("TauxSoudure")) Then
490                   dblTotalSoudure = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxSoudure"))
495                 Else
500                   dblTotalSoudure = 0
505                 End If
                    
510               Case "Assemblage":
515                 If Not IsNull(rstProjSoum.Fields("TauxAssemblage")) Then
520                   dblTotalAssemblage = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxAssemblage"))
525                 Else
530                   dblTotalAssemblage = 0
535                 End If
                    
540               Case "Peinture":
545                 If Not IsNull(rstProjSoum.Fields("TauxPeinture")) Then
550                   dblTotalPeinture = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxPeinture"))
555                 Else
560                   dblTotalPeinture = 0
565                 End If
                    
570               Case "Test":
575                 If Not IsNull(rstProjSoum.Fields("TauxTest")) Then
580                   dblTotalTest = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxTest"))
585                 Else
590                   dblTotalTest = 0
595                 End If
                    
600               Case "Installation":
605                 If Not IsNull(rstProjSoum.Fields("TauxInstallation")) Then
610                   dblTotalInstallation = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxInstallation"))
615                 Else
620                   dblTotalInstallation = 0
625                 End If
                    
630               Case "Formation":
635                 If Not IsNull(rstProjSoum.Fields("TauxFormation")) Then
640                   dblTotalFormation = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxFormation"))
645                 Else
650                   dblTotalFormation = 0
655                 End If
                    
660               Case "Gestion":
665                 If Not IsNull(rstProjSoum.Fields("TauxGestion")) Then
670                   dblTotalGestion = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxGestion"))
675                 Else
680                   dblTotalGestion = 0
685                 End If
                    
690               Case "Shipping":
695                 If Not IsNull(rstProjSoum.Fields("TauxShipping")) Then
700                   dblTotalShipping = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxShipping"))
705                 Else
710                   dblTotalShipping = 0
715                 End If
720             End Select
725           End If

730           Call rstPunch.MoveNext
735         Loop

740         Call rstPunch.Close
745         Set rstPunch = Nothing
750       Else
755         If Not IsNull(rstProjSoum.Fields("TempsDessin")) And Not IsNull(rstProjSoum.Fields("TauxDessin")) Then
760           dblTotalDessin = CDbl(rstProjSoum.Fields("TempsDessin")) * CDbl(rstProjSoum.Fields("TauxDessin"))
765         Else
770           dblTotalDessin = 0
775         End If

780         If Not IsNull(rstProjSoum.Fields("TempsCoupe")) And Not IsNull(rstProjSoum.Fields("TauxCoupe")) Then
785           dblTotalCoupe = CDbl(rstProjSoum.Fields("TempsCoupe")) * CDbl(rstProjSoum.Fields("TauxCoupe"))
790         Else
795           dblTotalCoupe = 0
800         End If

805         If Not IsNull(rstProjSoum.Fields("TempsMachinage")) And Not IsNull(rstProjSoum.Fields("TauxMachinage")) Then
810           dblTotalMachinage = CDbl(rstProjSoum.Fields("TempsMachinage")) * CDbl(rstProjSoum.Fields("TauxMachinage"))
815         Else
820           dblTotalMachinage = 0
825         End If

830         If Not IsNull(rstProjSoum.Fields("TempsSoudure")) And Not IsNull(rstProjSoum.Fields("TauxSoudure")) Then
835           dblTotalSoudure = CDbl(rstProjSoum.Fields("TempsSoudure")) * CDbl(rstProjSoum.Fields("TauxSoudure"))
840         Else
845           dblTotalSoudure = 0
850         End If

855         If Not IsNull(rstProjSoum.Fields("TempsAssemblage")) And Not IsNull(rstProjSoum.Fields("TauxAssemblage")) Then
860           dblTotalAssemblage = CDbl(rstProjSoum.Fields("TempsAssemblage")) * CDbl(rstProjSoum.Fields("TauxAssemblage"))
865         Else
870           dblTotalAssemblage = 0
875         End If

880         If Not IsNull(rstProjSoum.Fields("TempsPeinture")) And Not IsNull(rstProjSoum.Fields("TauxPeinture")) Then
885           dblTotalPeinture = CDbl(rstProjSoum.Fields("TempsPeinture")) * CDbl(rstProjSoum.Fields("TauxPeinture"))
890         Else
895           dblTotalPeinture = 0
900         End If

905         If Not IsNull(rstProjSoum.Fields("TempsTest")) And Not IsNull(rstProjSoum.Fields("TauxTest")) Then
910           dblTotalTest = CDbl(rstProjSoum.Fields("TempsTest")) * CDbl(rstProjSoum.Fields("TauxTest"))
915         Else
920           dblTotalTest = 0
925         End If

930         If Not IsNull(rstProjSoum.Fields("TempsInstallation")) And Not IsNull(rstProjSoum.Fields("TauxInstallation")) Then
935           dblTotalInstallation = CDbl(rstProjSoum.Fields("TempsInstallation")) * CDbl(rstProjSoum.Fields("TauxInstallation"))
940         Else
945           dblTotalInstallation = 0
950         End If

955         If Not IsNull(rstProjSoum.Fields("TempsFormation")) And Not IsNull(rstProjSoum.Fields("TauxFormation")) Then
960           dblTotalFormation = CDbl(rstProjSoum.Fields("TempsFormation")) * CDbl(rstProjSoum.Fields("TauxFormation"))
965         Else
970           dblTotalFormation = 0
975         End If
  
980         If Not IsNull(rstProjSoum.Fields("TempsGestion")) And Not IsNull(rstProjSoum.Fields("TauxGestion")) Then
985           dblTotalGestion = CDbl(rstProjSoum.Fields("TempsGestion")) * CDbl(rstProjSoum.Fields("TauxGestion"))
990         Else
995           dblTotalGestion = 0
1000        End If

1005        If Not IsNull(rstProjSoum.Fields("TempsShipping")) And Not IsNull(rstProjSoum.Fields("TauxShipping")) Then
1010          dblTotalShipping = CDbl(rstProjSoum.Fields("TempsShipping")) * CDbl(rstProjSoum.Fields("TauxShipping"))
1015        Else
1020          dblTotalShipping = 0
1025        End If
1030      End If
         
1035      dblTotalTemps = dblTotalDessin + _
                          dblTotalCoupe + _
                          dblTotalMachinage + _
                          dblTotalSoudure + _
                          dblTotalAssemblage + _
                          dblTotalPeinture + _
                          dblTotalTest + _
                          dblTotalInstallation + _
                          dblTotalFormation + _
                          dblTotalGestion + _
                          dblTotalShipping

1040      If m_eType = TYPE_PROJET Then
1045        dblHebergement = 0
1050        dblRepas = 0
1055        dblTransport = 0
1060        dblUniteMobile = 0
1065      Else
1070        If Not IsNull(rstProjSoum.Fields("TotalHebergement")) Then
1075          dblHebergement = CDbl(rstProjSoum.Fields("TotalHebergement"))
1080        Else
1085          dblHebergement = 0
1090        End If
          
1095        If Not IsNull(rstProjSoum.Fields("TotalRepas")) Then
1100          dblRepas = CDbl(rstProjSoum.Fields("TotalRepas"))
1105        Else
1110          dblRepas = 0
1115        End If

1120        If Not IsNull(rstProjSoum.Fields("TempsTransport")) Then
1125          dblTransport = CDbl(rstProjSoum.Fields("TempsTransport")) * CDbl(rstProjSoum.Fields("TauxTransport"))
1130        Else
1135          dblTransport = 0
1140        End If

1145        If Not IsNull(rstProjSoum.Fields("TempsUniteMobile")) Then
1150          dblUniteMobile = CDbl(rstProjSoum.Fields("TempsUniteMobile")) * CDbl(rstProjSoum.Fields("TauxUniteMobile"))
1155        Else
1160          dblUniteMobile = 0
1165        End If
1170      End If

          'Correction d'un bug de Type Incompatible
1175      If IsNumeric(rstProjSoum.Fields("PrixEmballage")) Then
1180        dblPrixEmballage = CDbl(rstProjSoum.Fields("PrixEmballage"))
1185      Else
1190        dblPrixEmballage = 0
1195      End If
      
1200      dblTotalResteTemps = dblHebergement + dblRepas + dblTransport + dblUniteMobile + dblPrixEmballage
                                                              
1205      If IsNumeric(rstProjSoum.Fields("total_manuel")) Then
1210        dblTotalManuel = CDbl(rstProjSoum.Fields("total_manuel"))
1215      Else
1220        dblTotalManuel = 0
1225      End If
                        
1230      dblTotalImprevue = Round((dblPrixPieces + dblProfit) * CDbl(rstProjSoum.Fields("Imprevue")), 2)
   
1235      dblPrixTotal = dblPrixPieces + dblProfit + dblTotalTemps + dblTotalImprevue + dblTotalManuel + dblTotalResteTemps
                          
          'Le calcul de la commission est sur les manuels (Nbre de page * prix par pages * nbre de copies) + (prix des pièces * pourcentage d'imprévus) + total des temps
1240      dblCommission = Round(dblPrixTotal * CDbl(rstProjSoum.Fields("Commission")), 2)
        
          'Le prix total est le calcul des manuels (Nbre de page * prix par pages * nbre de copies) + (prix des pièces * pourcentage d'imprévus) + total des temps + total de la commission
1245      dblGrandTotal = dblPrixTotal + dblCommission
                
          'Format monétaires avec 2 chiffres après la virgule
1250      rstProjSoum.Fields("Total_Piece") = Conversion(CStr(Round(dblPrixPieces, 2)), MODE_ARGENT)
1255      rstProjSoum.Fields("Total_Temps") = Conversion(CStr(Round(dblTotalTemps, 2)), MODE_ARGENT)
1260      rstProjSoum.Fields("PrixTotal") = Conversion(CStr(Round(dblGrandTotal, 2)), MODE_ARGENT)
1265      rstProjSoum.Fields("total_Imprevue") = Conversion(CStr(Round(dblTotalImprevue, 2)), MODE_ARGENT)
1270      rstProjSoum.Fields("total_commission") = Conversion(CStr(Round(dblCommission, 2)), MODE_ARGENT)
1275      rstProjSoum.Fields("total_profit") = Conversion(CStr(Round(dblProfit, 2)), MODE_ARGENT)

1280      Call rstProjSoum.Update

1285      Call rstPiece.Close
1290      Set rstPiece = Nothing
1295    Else
1300      If m_eType = TYPE_PROJET Then
1305        Call MsgBox("Le projet " & sNoProjSoum & " est inexistant!", vbOKOnly, "Erreur")
1310      Else
1315        Call MsgBox("La soumission " & sNoProjSoum & " est inexistant!", vbOKOnly, "Erreur")
1320      End If
1325    End If

1330    Call rstProjSoum.Close
1335    Set rstProjSoum = Nothing

1340    Exit Sub

AfficherErreur:

1345    woups "frmProjSoumMec", "CalculerTotalRecordset", Err, Erl
End Sub

Private Sub ChoisirFournisseurMateriel()

5       On Error GoTo AfficherErreur

        'On ajoute la pièce en négatif dans le listview
10      Dim rstProjet  As ADODB.Recordset
15      Dim rstConfig  As ADODB.Recordset
20      Dim itmAncien  As ListItem
25      Dim itmNouveau As ListItem
30      Dim sQuantite  As String
35      Dim sExtra     As String
40      Dim sTauxUSA   As String
45      Dim sTauxSPA   As String

50      Set rstConfig = New ADODB.Recordset

55      Call rstConfig.Open("SELECT TauxAmericain, TauxEspagnol FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)
  
60      sTauxUSA = rstConfig.Fields("TauxAmericain")
65      sTauxSPA = rstConfig.Fields("TauxEspagnol")

70      Call rstConfig.Close
75      Set rstConfig = Nothing
  
80      If m_bChangementFRS = True Then
85        If lvwfournisseur.SelectedItem.Text <> "CHOISIR ULTÉRIEUREMENT" Then
90          lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_DISTRIB) = lvwfournisseur.SelectedItem.Text
95          lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_DISTRIB).Tag = lvwfournisseur.SelectedItem.Tag

100         If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor <> COLOR_BRUN Then
              'Prix listé
105           If Trim$(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_LIST)) = vbNullString Then
110             lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion("0", MODE_ARGENT, 4)
115           Else
120             If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "USA" Then
125               lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_LIST)) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
130             Else
135               If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "SPA" Then
140                 lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_LIST)) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
145               Else
150                 lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_LIST), MODE_ARGENT, 4)
155               End If
160             End If
165           End If
       
170           lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PRIX_LIST).Tag
       
              'Si il y a un prix net, on le met l'escompte et le prix net sinon, on prend le prix
              'spécial pour mettre dans le prix net
175           If Trim$(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_NET)) <> vbNullString Then
180             lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_ESCOMPTE), MODE_POURCENT)

185             If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "USA" Then
190               lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_NET)) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
195             Else
200               If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "SPA" Then
205                 lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_NET)) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
210               Else
215                 lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_NET), MODE_ARGENT, 4)
220               End If
225             End If
230           Else
235             If Trim$(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_SP)) <> vbNullString Then
240               If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "USA" Then
245                 lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_SP)) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
250               Else
255                 If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "SPA" Then
260                   lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_SP)) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
265                 Else
270                   lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_SP), MODE_ARGENT, 4)
275                 End If
280               End If
285             Else
290               lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion("0", MODE_POURCENT)
295               lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PRIX_NET) = Conversion("0", MODE_ARGENT, 4)
300             End If
305           End If
              
310           lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_TOTAL) = Conversion(Round(Replace(lvwSoumission.SelectedItem.Text, "*", vbNullString) * lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PRIX_NET) * CSng(m_sProfit), 2), MODE_ARGENT, 2)

315           lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_TOTAL).Tag = lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag
      
              'Pour le profit, c'est le prix total - (prix net * quantité)
320           lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PROFIT) = Conversion(CStr(Round(lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_TOTAL) - (lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PRIX_NET) * Replace(lvwSoumission.SelectedItem.Text, "*", vbNullString)), 2)), MODE_ARGENT)
325         End If
330       Else
335         lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_DISTRIB) = vbNullString
340         lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_DISTRIB).Tag = 0

345         lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion("0", MODE_ARGENT, 2)
350         lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion("0", MODE_POURCENT)
355         lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PRIX_NET) = Conversion("0", MODE_ARGENT, 2)

360         lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_TOTAL) = Conversion(Round(Replace(lvwSoumission.SelectedItem.Text, "*", vbNullString) * lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PRIX_NET) * CSng(m_sProfit), 2), MODE_ARGENT, 2)
        
365         lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PROFIT) = Conversion(CStr(Round(lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_TOTAL) - (lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PRIX_NET) * Replace(lvwSoumission.SelectedItem.Text, "*", vbNullString)), 2)), MODE_ARGENT)

370         If m_eType = TYPE_PROJET Then
375           If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_COMMENTAIRE) <> "" Then
380             lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_COMMENTAIRE).ForeColor = COLOR_MAGENTA
385           End If

390           If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_FACTURATION) <> "" Then
395             lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_FACTURATION).ForeColor = COLOR_MAGENTA
400           End If
405         End If

410         lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_DESCR).ForeColor = COLOR_MAGENTA
415         lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = COLOR_MAGENTA
420         lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = COLOR_MAGENTA
425         lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_MAGENTA
430         lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = COLOR_MAGENTA
435         lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = COLOR_MAGENTA
440         lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = COLOR_MAGENTA
445         lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = COLOR_MAGENTA

450         Call lvwSoumission.Refresh
455       End If

460       Call CalculerPrix

          'On cache le listview
465       frafournisseur.Visible = False

470       m_bPieceInutile = False
475       m_bChangementFRS = False
480     Else
485       If Right$(txtNoProjSoum.Text, 2) >= "01" And Right$(txtNoProjSoum.Text, 2) <= "19" Then
490         sExtra = InputBox("Dans quel extra le retour doit être fait ? (2 chiffres seulement)")

495         If Len(sExtra) <> 2 Then
500           Call MsgBox("Format incorrect!", vbOKOnly, "Erreur")

505           Exit Sub
510         End If

515         If Not IsNumeric(sExtra) Then
520           Call MsgBox("L'extra doit être numérique!", vbOKOnly, "Erreur")

525           Exit Sub
530         End If

535         If sExtra < 60 Or sExtra > 98 Then
540           Call MsgBox("L'extra doit être entre 60 et 98!", vbOKOnly, "Erreur")

545           Exit Sub
550         End If

555         Set rstProjet = New ADODB.Recordset

560         Call rstProjet.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sExtra & "'", g_connData, adOpenDynamic, adLockOptimistic)

565         If rstProjet.EOF Then
570           Call MsgBox("Le projet " & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sExtra & " n'existe pas!", vbOKOnly, "Erreur")

575           Call rstProjet.Close
580           Set rstProjet = Nothing

585           Exit Sub
590         Else
595           Call rstProjet.Close
600           Set rstProjet = Nothing
605         End If
610       End If
          
          'Saisie de la quantité
615       sQuantite = InputBox("Quelle est la quantité?")

620       sQuantite = Replace(sQuantite, ".", ",")

625       sQuantite = Replace(sQuantite, "-", "")
    
630       If sQuantite <> vbNullString Then
635         If Not IsNumeric(sQuantite) Then
640           Call MsgBox("Quantité non numérique!", vbOKOnly, "Erreur")
      
645           Exit Sub
650         End If
655       Else
660         Exit Sub
665       End If

670       If CDbl(sQuantite) <= CDbl(Replace(lvwSoumission.SelectedItem.Text, "*", vbNullString)) Then
675         Set itmAncien = lvwSoumission.SelectedItem
680         Set itmNouveau = lvwSoumission.ListItems.Add(itmAncien.Index + 1)

685         itmNouveau.Checked = itmAncien.Checked
  
            'Quantité
690         itmNouveau.Text = "-" & sQuantite

            'On met l'id de la section dans le tag du listItem
695         itmNouveau.Tag = itmAncien.Tag
                                                                                                         
            'No d'item
700         itmNouveau.SubItems(I_COL_SOUM_PIECE) = itmAncien.SubItems(I_COL_SOUM_PIECE)
  
            'On met le nom de la sous-section dans le tag du no d'item
705         itmNouveau.ListSubItems(I_COL_SOUM_PIECE).Tag = itmAncien.ListSubItems(I_COL_SOUM_PIECE).Tag
  
            'On met la description en francais dans la colonne et la description en anglais
            'dans le tag
710         itmNouveau.SubItems(I_COL_SOUM_DESCR) = itmAncien.SubItems(I_COL_SOUM_DESCR)
715         itmNouveau.ListSubItems(I_COL_SOUM_DESCR).Tag = itmAncien.ListSubItems(I_COL_SOUM_DESCR).Tag
          
            'On met le nom du fabricant dans la colonne et l'ordre de la section dans le tag
720         itmNouveau.SubItems(I_COL_SOUM_MANUFACT) = itmAncien.SubItems(I_COL_SOUM_MANUFACT)
725         itmNouveau.ListSubItems(I_COL_SOUM_MANUFACT).Tag = itmAncien.ListSubItems(I_COL_SOUM_MANUFACT).Tag

            'Prix listé
730         itmNouveau.SubItems(I_COL_SOUM_PRIX_LIST) = itmAncien.SubItems(I_COL_SOUM_PRIX_LIST)

735         itmNouveau.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = itmAncien.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag
       
740         itmNouveau.SubItems(I_COL_SOUM_ESCOMPTE) = itmAncien.SubItems(I_COL_SOUM_ESCOMPTE)

745         itmNouveau.SubItems(I_COL_SOUM_PRIX_NET) = itmAncien.SubItems(I_COL_SOUM_PRIX_NET)
            
            'On met le fournisseur dans la colonne et l'id dans le tag
750         itmNouveau.SubItems(I_COL_SOUM_DISTRIB) = lvwfournisseur.SelectedItem.Text
755         itmNouveau.ListSubItems(I_COL_SOUM_DISTRIB).Tag = lvwfournisseur.SelectedItem.Tag
    
            'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
760         itmNouveau.SubItems(I_COL_SOUM_TOTAL) = Conversion(Round(CDbl(Replace(itmNouveau.Text, "*", "") * CDbl(itmNouveau.SubItems(I_COL_SOUM_PRIX_NET)) * CDbl(m_sProfit)), 2), MODE_ARGENT)
      
            'Pour le profit, c'est le prix total - (prix net * quantité)
765         itmNouveau.SubItems(I_COL_SOUM_PROFIT) = Conversion(Round(CDbl(itmNouveau.SubItems(I_COL_SOUM_TOTAL)) - (CDbl(Replace(itmNouveau.Text, "*", "") * CDbl(itmNouveau.SubItems(I_COL_SOUM_PRIX_NET)))), 2), MODE_ARGENT)

770         If Right$(txtNoProjSoum.Text, 2) >= "01" And Right$(txtNoProjSoum.Text, 2) <= "19" Then
              'Pour savoir lors de l'enregistrement qu'il faut le lier avec le 81
775           itmNouveau.ListSubItems(I_COL_SOUM_PROFIT).Tag = "RETOUR " & sExtra
780         End If

785         itmNouveau.SubItems(I_COL_SOUM_DATE_COMMANDE) = " "
790         itmNouveau.SubItems(I_COL_SOUM_DATE_REQUISE) = " "
795         itmNouveau.SubItems(I_COL_SOUM_NOM_COMMANDE) = " "
800         itmNouveau.SubItems(I_COL_SOUM_NO_SEQUENTIEL) = " "

805         If itmAncien.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_NOIR Then
810           itmNouveau.ListSubItems(I_COL_SOUM_DATE_COMMANDE).ForeColor = COLOR_NOIR
815           itmNouveau.ListSubItems(I_COL_SOUM_DATE_REQUISE).ForeColor = COLOR_NOIR
820           itmNouveau.ListSubItems(I_COL_SOUM_DESCR).ForeColor = COLOR_NOIR
825           itmNouveau.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = COLOR_NOIR
830           itmNouveau.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = COLOR_NOIR
835           itmNouveau.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = COLOR_NOIR
840           itmNouveau.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_NOIR
845           itmNouveau.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = COLOR_NOIR
850           itmNouveau.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = COLOR_NOIR
855           itmNouveau.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = COLOR_NOIR
860           itmNouveau.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = COLOR_NOIR
865         Else
870           itmNouveau.ListSubItems(I_COL_SOUM_DATE_COMMANDE).ForeColor = COLOR_BRUN
875           itmNouveau.ListSubItems(I_COL_SOUM_DATE_REQUISE).ForeColor = COLOR_BRUN
880           itmNouveau.ListSubItems(I_COL_SOUM_DESCR).ForeColor = COLOR_BRUN
885           itmNouveau.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = COLOR_BRUN
890           itmNouveau.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = COLOR_BRUN
895           itmNouveau.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = COLOR_BRUN
900           itmNouveau.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_BRUN
905           itmNouveau.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = COLOR_BRUN
910           itmNouveau.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = COLOR_BRUN
915           itmNouveau.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = COLOR_BRUN
920           itmNouveau.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = COLOR_BRUN
925         End If

            'Calcul des prix
930         Call CalculerPrix
  
            'On cache le listview
935         frafournisseur.Visible = False

940         m_bPieceInutile = False

            'Resélectionne le premier élément du listview
945         If lvwSoumission.ListItems.count > 0 Then
950           Call Deselect

955           lvwSoumission.ListItems(1).Selected = True
960         End If
965       Else
970         Call MsgBox("Quantité trop grande!", vbOKOnly, "Erreur")
975       End If
980     End If

985     Exit Sub

AfficherErreur:

990     woups "frmProjSoumMec", "ChoisirFournisseurMateriel", Err, Erl
End Sub

Private Sub ChoisirFournisseur()
  
5       On Error GoTo AfficherErreur

        'On ajoute la pièce dans lvwSoumission
10      Dim sQuantite    As String
15      Dim sSousSection As String
20      Dim bDemanderSS  As Boolean
        Dim sParams      As String
        
        'Si l'utilisateur a déjà choisi un emplacement, il ne faut pas
        'lui demander dans quelle sous-section
  
        'Si il y a des enregistrements dans le listview
25      If lvwSoumission.ListItems.count > 0 Then
          'Si le premier n'est pas sélectionné.. celui-ci est sélectionné par défaut
30        If lvwSoumission.SelectedItem.Index > 1 Then
            'Si l'emplacement est valide
35          If VerifierEmplacement(lvwSoumission.SelectedItem.Index) = True Then
              'Si c'est une sous-section
40            If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).Tag = vbNullString Then
                'Si l'autre d'au dessus est une section
45              If lvwSoumission.ListItems(lvwSoumission.SelectedItem.Index - 1).Tag = vbNullString Then
                  'Message d'erreur
50                Call MsgBox("Vous ne pouvez pas mettre une pièce entre une section et une sous-section", vbOKOnly, "Erreur")
          
55                frafournisseur.Visible = False
          
                  'Il faut resélectionné le premier pour faire comme si il n'était plus
                  'sélectionné
60                Call Deselect
                  
65                lvwSoumission.ListItems(1).Selected = True
          
70                Exit Sub
75              Else
                  'Sinon, on prend le tag de la section d'en haut
80                sSousSection = lvwSoumission.ListItems(lvwSoumission.SelectedItem.Index - 1).ListSubItems(I_COL_SOUM_PIECE).Tag
85              End If
90            Else
                'On prend le tag de l'élément sélectionné
95              sSousSection = lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).Tag
100           End If
105         Else
110           If lvwSoumission.ListItems(lvwSoumission.SelectedItem.Index - 1).ListSubItems(I_COL_SOUM_MANUFACT).Tag <> "" Then
115             If MsgBox("Vous essayez d'ajouter une pièce de la section " & cmbSections.Text & " dans la section " & cmbSections.LIST(lvwSoumission.ListItems(lvwSoumission.SelectedItem.Index - 1).ListSubItems(I_COL_SOUM_MANUFACT).Tag - 1) & vbNewLine & "Voulez-vous ajouter la pièce dans la section " & cmbSections.LIST(lvwSoumission.ListItems(lvwSoumission.SelectedItem.Index - 1).ListSubItems(I_COL_SOUM_MANUFACT).Tag - 1), vbYesNo, "Erreur") = vbYes Then
120               cmbSections.ListIndex = lvwSoumission.ListItems(lvwSoumission.SelectedItem.Index - 1).ListSubItems(I_COL_SOUM_MANUFACT).Tag - 1

125               Call ChoisirFournisseur
130             End If
        
135             frafournisseur.Visible = False
        
                'Il faut resélectionné le premier pour faire comme si il n'était plus
                'sélectionné
140             Call Deselect
              
145             lvwSoumission.ListItems(1).Selected = True
        
150             Exit Sub
155           Else
160             Call MsgBox("Impossible d'ajouter entre une section et une sous-section!", vbOKOnly, "Erreur")

165             Exit Sub
170           End If
175         End If
180       Else
185         bDemanderSS = True
190       End If
195     Else
          'Sinon, on demande la section
200       bDemanderSS = True
205     End If
  
        'Saisie de la quantité
210     sQuantite = InputBox("Quelle est la quantité?")

215     sQuantite = Replace(sQuantite, ".", ",")
    
220     If sQuantite <> vbNullString Then
225       If Not IsNumeric(sQuantite) Then
230         Call MsgBox("Quantité non numérique!", vbOKOnly, "Erreur")
      
235         Exit Sub
240       Else
245         If sQuantite < 0 Then
250           If lvwfournisseur.SelectedItem.Text = "CHOISIR ULTÉRIEUREMENT" Then
255             Call MsgBox("Impossible de faire une demande de prix sur une pièce négative!", vbOKOnly, "Erreur")

260             Exit Sub
265           End If
270         End If
275       End If
280     Else
285       Exit Sub
290     End If

295     If bDemanderSS = True Then
300       If m_sSousSection <> S_PAS_SOUS_SECTION Then
305         sSousSection = InputBox("Quelle est la sous-section?", , m_sSousSection)
310       Else
315         sSousSection = InputBox("Quelle est la sous-section?")
320       End If
325     End If
    
        'Si la sous-section est vide
330     If sSousSection = vbNullString Then
          'On initialise la sous-section à "PAS DE SOUS-SECTIONS"
335       sSousSection = S_PAS_SOUS_SECTION
340       m_sSousSection = vbNullString
345     Else
350       m_sSousSection = sSousSection
355     End If
  
360     If sQuantite < 0 Then
365       If m_eType = TYPE_PROJET Then
370         If Right$(txtNoProjSoum.Text, 2) >= 60 And Right$(txtNoProjSoum.Text, 2) <= 98 Then
375           Call AjouterNegatifDansListView(CDbl(sQuantite), sSousSection)
380         Else
385           Call AjouterDansListViewSoumission(CDbl(sQuantite), sSousSection)
390         End If
395       Else
400         Call AjouterDansListViewSoumission(CDbl(sQuantite), sSousSection)
405       End If
410     Else
415       Call AjouterDansListViewSoumission(CDbl(sQuantite), sSousSection)
420     End If
  
        'Calcul des prix
425     Call CalculerPrix
  
        'On cache le listview
430     frafournisseur.Visible = False
  
        'Resélectionne le premier élément du listview
435     If lvwSoumission.ListItems.count > 0 Then
440       Call Deselect

445       lvwSoumission.ListItems(1).Selected = True
450     End If

455     Exit Sub

AfficherErreur:

If Err.number = 13 And Erl = 115 Then
  sParams = "cmbSections.Text : " & cmbSections.Text & "   " & _
            "No Proj/Soum : " & txtNoProjSoum.Text & "   " & _
            "lvwSoumission.SelectedItem.Index - 1 : " & lvwSoumission.SelectedItem.Index - 1 & "   " & _
            "lvwSoumission.ListItems(lvwSoumission.SelectedItem.Index - 1).ListSubItems(I_COL_SOUM_MANUFACT).Tag : " & lvwSoumission.ListItems(lvwSoumission.SelectedItem.Index - 1).ListSubItems(I_COL_SOUM_MANUFACT).Tag
            
  woups "frmProjSoumMec", "ChoisirFournisseur", Err, Erl, sParams)
Else
460     woups "frmProjSoumMec", "ChoisirFournisseur", Err, Erl
End If
End Sub

Private Sub lvwFournisseur_DblClick()

5       On Error GoTo AfficherErreur

10      If m_bPieceInutile = True Or m_bChangementFRS = True Then
15        Call ChoisirFournisseurMateriel
20      Else
25        Call ChoisirFournisseur
30      End If

35      Exit Sub

AfficherErreur:

40      woups "frmProjSoumMec", "lvwFournisseur_DblClick", Err, Erl
End Sub

Private Sub lvwPieces_DblClick()

5       On Error GoTo AfficherErreur

10      m_bPieceInutile = False
15      m_bRecherchePiece = False
20      m_bChangementFRS = False

        'On affiche lvwFournisseur selon la pièce choisie
        'Rempli les fournisseurs de la pièce choisie
25      If cmbSections.ListCount > 0 Then
30        Call AfficherListeFournisseurs

          'si le listview n'est pas vide
35        If lvwfournisseur.ListItems.count = 1 Then
40          If MsgBox("Il n'y a aucun fournisseur pour cette pièce!" & vbNewLine & "Voulez-vous en ajouter?", vbYesNo, "Erreur") = vbYes Then
45            Screen.MousePointer = vbHourglass

              'On ouvre le catalogue sur cet enregistrement
50            Call FrmCatalogueMec.AfficherForm(cmbPieces.Text, lvwPieces.SelectedItem.SubItems(I_COL_PIECES_MANUFACT), lvwPieces.SelectedItem.SubItems(I_COL_PIECES_NO_ITEM))

55            Screen.MousePointer = vbDefault
60          End If
65        End If
70      Else
75        Call MsgBox("Il n'y a pas de sections!", vbOKOnly, "Erreur")
80      End If

85      Exit Sub

AfficherErreur:

90      woups "frmProjSoumMec", "lvwPieces_DblClick", Err, Erl
End Sub

Private Sub AfficherListeFournisseurs()

5       On Error GoTo AfficherErreur

        'Méthode qui sert à affiche la liste des fournisseurs
        'Affiche le frame seulement si il y a des items dans listview
10      Call RemplirListViewFournisseur

15      If lvwfournisseur.ListItems.count > 1 Then
20        If m_bRecherchePiece = True Then
25          fraPieceTrouve.Visible = False
30        End If

35        frafournisseur.Visible = True
40        Call lvwfournisseur.SetFocus
45      End If

50      Exit Sub

AfficherErreur:

55      woups "frmProjSoumMec", "AfficherListeFournisseurs", Err, Erl
End Sub

Private Sub lvwPieces_KeyDown(KeyCode As Integer, Shift As Integer)

5       On Error GoTo AfficherErreur

10      Dim sTexte As String

15      If Shift = vbCtrlMask Then
20        If KeyCode = vbKeyF Then
25          sTexte = InputBox("Quel est le texte à rechercher?")

30          If Trim$(sTexte) <> vbNullString Then
35            If Len(Trim$(sTexte)) >= 2 Then
40              Call RemplirListViewRecherche(I_COL_PIECES_NO_ITEM, sTexte)

45              If lvwPieceTrouve.ListItems.count > 0 Then
50                fraPieceTrouve.Visible = True
55              Else
60                Call MsgBox("Aucun enregistrements trouvés!", vbOKOnly, "Erreur")
65              End If
70            Else
75              Call MsgBox("Il faut un minimum de 2 caractères pour rechercher!", vbOKOnly, "Erreur")
80            End If
85          End If
90        End If
95      Else
100       If KeyCode = vbKeyReturn Then
105         Call lvwPieces_DblClick
110       End If
115     End If

120     Exit Sub

AfficherErreur:

125     woups "frmProjSoumMec", "lvwPieces_KeyDown", Err, Erl
End Sub

Private Sub lvwSoumission_DblClick()

5       On Error GoTo AfficherErreur

        'Si il y a des enregistrements
10      If lvwSoumission.ListItems.count > 0 Then
15        If m_eMode = MODE_AJOUT_MODIF Then
20          If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) = vbNullString Then
25            Call ModifierSousSection
30          Else
              'Si c'est une pièce
35            If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) <> vbNullString And lvwSoumission.SelectedItem.Tag <> vbNullString Then
40              If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) <> "Texte" And lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) <> "Text" Then
                  'Si la pièce n'a pas de fournisseur
45                If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_DISTRIB) = vbNullString Then
50                  If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PROFIT).Tag <> "EXTRA" Then
55                    Call AjouterPrix
60                  Else
65                    Call MsgBox("Cette commande doit être faite dans le projet " & lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PROVENANCE), vbOKOnly, "Erreur")
70                  End If
75                Else
                    'Si le listItem est orange
80                  If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_ORANGE Then
85                    If MsgBox("Voulez-vous annuler cette commande?", vbYesNo) = vbYes Then
90                      Call AnnulerCommande
95                    End If
100                 Else
105                   If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_BRUN Then
110                     Call ChangerFournisseurRetour
115                   End If
120                 End If
125               End If
130             Else
135               Call ModifierTexte
140             End If
145           End If
150         End If
155       End If
160     End If

165     Exit Sub

AfficherErreur:

170     woups "frmProjSoumMec", "lvwSoumission_DblClick", Err, Erl
End Sub

Private Sub AjouterPrix()

5       On Error GoTo AfficherErreur

10      Call ViderChamps_frs

        'Rempli le combo des fournisseurs
15      Call RemplirComboFournisseur

20      cmbfrs.Locked = False

25      m_bMauvaisPrix = False

        'Positionne le frame
30      fraPrix.Top = lvwSoumission.Top + 500

        'Montre le frame
35      fraPrix.Visible = True

        'Met le numéro de la pièce dans le tag
40      fraPrix.Tag = lvwSoumission.SelectedItem.Index

        'Donne le focus au combo
45      Call cmbfrs.SetFocus

50      Exit Sub

AfficherErreur:

55      woups "frmProjSoumMec", "AjouterPrix", Err, Erl
End Sub

Private Sub ModifierTexte()

5       On Error GoTo AfficherErreur

10      Dim sTexte As String

15      sTexte = InputBox("Quel est le nouveau texte?", , lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_DESCR))

20      If sTexte <> "" Then
25        If Len(sTexte) > 255 Then
30          Call MsgBox("Le texte ne doit pas dépasser 255 caractères!", vbOKOnly, "Erreur")
35        Else
40          lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_DESCR) = sTexte
45        End If
50      End If

55      Exit Sub

AfficherErreur:

60      woups "frmProjSoumMec", "ModifierTexte", Err, Erl
End Sub

Private Sub ModifierSousSection()

5       On Error GoTo AfficherErreur

10      Dim sSousSection As String
15      Dim sAncienneSS  As String
20      Dim sTag         As String
25      Dim iCompteur    As Integer

30      sSousSection = InputBox("Quel est le nouveau nom de la sous-section?", , lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_DESCR))

35      If StrPtr(sSousSection) <> 0 Then
40        sAncienneSS = lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_DESCR)
        
45        If sAncienneSS = vbNullString Then
50          sAncienneSS = S_PAS_SOUS_SECTION
55        End If
        
60        If Trim$(sSousSection) = vbNullString Then
65          sTag = S_PAS_SOUS_SECTION
70          sSousSection = vbNullString
75        Else
80          sTag = sSousSection
85        End If

90        lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_DESCR) = sSousSection
95        lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_DESCR).Bold = True
      
100       For iCompteur = lvwSoumission.SelectedItem.Index + 1 To lvwSoumission.ListItems.count
105         If lvwSoumission.ListItems(iCompteur).ListSubItems(I_COL_SOUM_PIECE).Tag = sAncienneSS Then
110           lvwSoumission.ListItems(iCompteur).ListSubItems(I_COL_SOUM_PIECE).Tag = sTag
115         Else
120           Exit For
125         End If
130       Next
135     End If

140     Exit Sub

AfficherErreur:

145     woups "frmProjSoumMec", "ModifierSousSection", Err, Erl
End Sub

Private Sub ChangerFournisseurRetour()

5       On Error GoTo AfficherErreur

10      m_bRecherchePiece = False
15      m_bChangementFRS = True

20      Call AfficherListeFournisseurs

25      If lvwfournisseur.ListItems.count = 0 Then
30        Call MsgBox("Il n'y a aucun fournisseur pour cette pièce!", vbOKOnly, "Erreur")
 
35        Exit Sub
40      Else
45        frafournisseur.Visible = True
50      End If

55      Exit Sub

AfficherErreur:

60      woups "frmProjSoumMec", "ChangerFournisseurRetour", Err, Erl
End Sub


Private Sub RechercherPieceListViewSoumission(ByVal sTexte As String)

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer
15      Dim iSelected As Integer
20      Dim bTrouve   As Boolean

25      If lvwSoumission.SelectedItem.Index = 1 Then
30        iSelected = 1
35      Else
40        If lvwSoumission.SelectedItem.Index + 1 > lvwSoumission.ListItems.count Then
45          iSelected = 1
50        Else
55          iSelected = lvwSoumission.SelectedItem.Index + 1
60        End If
65      End If

70      For iCompteur = iSelected To lvwSoumission.ListItems.count
75        If InStr(1, UCase(lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE)), UCase(sTexte)) > 0 Then
80          Call lvwSoumission.SetFocus

81          Call Deselect

85          lvwSoumission.ListItems(iCompteur).Selected = True

90          Call lvwSoumission.SelectedItem.EnsureVisible

95          bTrouve = True

100         Exit For
105       End If
110     Next

115     If bTrouve = False Then
120       For iCompteur = 1 To iSelected - 1
125         If InStr(1, UCase(lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE)), UCase(sTexte)) > 0 Then
130           Call lvwSoumission.SetFocus

131           Call Deselect

135           lvwSoumission.ListItems(iCompteur).Selected = True

140           Call lvwSoumission.SelectedItem.EnsureVisible

145           bTrouve = True

150           Exit For
155         End If
160       Next
165     End If

170     If bTrouve = False Then
175       Call MsgBox("Aucun enregistrement trouvé!", vbOKOnly, "Erreur")
180     End If

185     Exit Sub

AfficherErreur:

190     woups "frmProjSoumMec", "RechercherPieceListViewSoumission", Err, Erl
End Sub

Private Sub lvwSoumission_ItemCheck(ByVal Item As MSComctlLib.ListItem)

5       On Error GoTo AfficherErreur

        'Si le programme n'est pas en mode ajout ou modif
10      If m_eMode = MODE_INACTIF Then
          'On annule le ItemCheck
15        Item.Checked = Not Item.Checked
20      Else
          'Si c'est une section, une sous-section ou du texte
25        If Item.Tag = vbNullString Or Item.SubItems(I_COL_SOUM_PIECE) = vbNullString Then
            'On annule le ItemCheck
30          Item.Checked = Not Item.Checked
35        End If
40      End If

45      Exit Sub

AfficherErreur:

50      woups "frmProjSoumMec", "lvwSoumission_ItemCheck", Err, Erl
End Sub

Private Sub lvwSoumission_KeyDown(KeyCode As Integer, Shift As Integer)

5       On Error GoTo AfficherErreur

        'Si le ListView n'est pas vide
10      If lvwSoumission.ListItems.count > 0 Then
15        If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor = -2147483640 Then
20          lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor = vbBlack
25        End If

30        If Shift = vbCtrlMask Then
35          If KeyCode = vbKeyF Then
40            m_sTexteRecherche = InputBox("Quel est la pièce à rechercher?")

45            If Trim$(m_sTexteRecherche) <> vbNullString Then
50              Call RechercherPieceListViewSoumission(m_sTexteRecherche)
55            End If
60          Else
65            If KeyCode = vbKeyN Then
70              If Trim$(m_sTexteRecherche) <> vbNullString Then
75                Call RechercherPieceListViewSoumission(m_sTexteRecherche)
80              End If
85            Else
                'S'il n'est pas en mode affichage
90              If m_bModeAffichage = False Then
                  'Si ce n'est pas une sous-section
95                If KeyCode = vbKeyC Then
100                 Call CopierPiece
105               Else
110                 If KeyCode = vbKeyV Then
115                   Call CollerPiece
120                 End If
125               End If
130             End If
135           End If
140         End If
145       Else
            'S'il n'est pas en mode affichage
150         If m_bModeAffichage = False Then
              'Si ce n'est pas une section
155           If lvwSoumission.SelectedItem.Tag <> vbNullString Then
                'Si ce n'est pas une sous-section
160             If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) <> vbNullString Then
                  'Si la touche pesée est Delete
165               If KeyCode = vbKeyDelete Then
                    'On l'efface
170                 Call EffacerItemListViewSoumission
175               Else
180                 If KeyCode = vbKeyReturn Or KeyCode = vbKeyN Then
185                   If m_eType = TYPE_PROJET Then
190                     If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) <> "Texte" And lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) <> "Text" Then
195                       If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PROFIT).Tag <> "EXTRA" Then
200                         If KeyCode = vbKeyReturn Then
205                           Call FacturerDate
210                         Else
215                           Call FacturerNC
220                         End If
225                       Else
230                         Call MsgBox("Cette commande doit être faite dans le projet " & lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PROVENANCE), vbOKOnly, "Erreur")
235                       End If
240                     End If
245                   End If
250                 End If
255               End If
260             End If
265           End If
270         End If
275       End If
280     End If

285     Exit Sub

AfficherErreur:

290     woups "frmProjSoumMec", "lvwSoumission_KeyDown", Err, Erl
End Sub

Private Sub FacturerDate()

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer

15      For iCompteur = 1 To lvwSoumission.ListItems.count
          'Si c'est pas une section
20        If lvwSoumission.ListItems(iCompteur).Tag <> "" Then
            'Si c'est pas une sous-section
25          If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> "" Then
              'Si la pièce est sélectionnée
30            If lvwSoumission.ListItems(iCompteur).Selected = True Then
                'Si il y a une date d'écrit
35              If Left$(lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_FACTURATION), 2) = "F-" Then
                  'On l'enlève
40                lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_FACTURATION) = ""
45              Else
                  'Si il n'y a rien
50                If Trim$(lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_FACTURATION)) = "" Then
                    'On ajoute la date de facturation
55                  lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_FACTURATION) = "F-" & txtDateFacturation.Text
60                End If
65              End If
70            End If
75          End If
80        End If
85      Next

90      Exit Sub

AfficherErreur:

95      woups "frmProjSoumMec", "FacturerDate", Err, Erl
End Sub

Private Sub FacturerNC()

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer

15      For iCompteur = 1 To lvwSoumission.ListItems.count
          'Si c'est pas une section
20        If lvwSoumission.ListItems(iCompteur).Tag <> "" Then
            'Si c'est pas une sous-section
25          If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> "" Then
              'Si la pièce est sélectionnée
30            If lvwSoumission.ListItems(iCompteur).Selected = True Then
                'Si il y a NC d'écrit
35              If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_FACTURATION) = "NC" Then
                  'On l'enlève
40                lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_FACTURATION) = ""
45              Else
                  'Si il n'y a rien
50                If Trim$(lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_FACTURATION)) = "" Then
                    'On ajoute la date de facturation
55                  lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_FACTURATION) = "NC"
60                End If
65              End If
70            End If
75          End If
80        End If
85      Next

90      Exit Sub

AfficherErreur:

95      woups "frmProjSoumMec", "FacturerNC", Err, Erl
End Sub

Private Sub EffacerItemListViewSoumission()

5       On Error GoTo AfficherErreur

10      Dim bSeulSS       As Boolean 'Pour savoir si c'est le seul enr. dans la sous-section
15      Dim bSeulS        As Boolean 'Pour savoir si c'est le seul enr. dans la section
20      Dim iIndex        As Integer
25      Dim itmPrecedent  As ListItem
30      Dim itmSuivant    As ListItem
35      Dim iCompteur     As Integer
40      Dim sMessage      As String
45      Dim iNbreSelected As Integer
50      Dim bSupprimer    As Boolean
55      Dim bPermission   As Boolean

60      For iCompteur = 1 To lvwSoumission.ListItems.count
65        If lvwSoumission.ListItems(iCompteur).Selected = True Then
70          iNbreSelected = iNbreSelected + 1

75          If iNbreSelected > 1 Then
80            Exit For
85          End If
90        End If
95      Next

100     If iNbreSelected > 1 Then
105       sMessage = "Voulez-vous vraiment effacer ces pièces?"
110     Else
115       sMessage = "Voulez-vous vraiment effacer cette pièce?"
120     End If

125     If m_eType = TYPE_SOUMISSION Then
130       bPermission = True
135     Else
140       If iNbreSelected > 1 Then
145         bPermission = True
150       Else
155         For iCompteur = 1 To lvwSoumission.ListItems.count
160           If lvwSoumission.ListItems(iCompteur).Selected = True Then
165             Exit For
170           End If
175         Next

180         If lvwSoumission.ListItems(iCompteur).ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_NOIR Or _
               lvwSoumission.ListItems(iCompteur).ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_BRUN Or _
               lvwSoumission.ListItems(iCompteur).ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_MAGENTA Then
185           bPermission = True
190         End If
195       End If
200     End If

205     iCompteur = 1

210     If bPermission = True Then
215       If MsgBox(sMessage, vbYesNo) = vbYes Then
220         Do While iCompteur <= lvwSoumission.ListItems.count
225           bSupprimer = False
230           bSeulS = False
235           bSeulSS = False

240           If lvwSoumission.ListItems(iCompteur).Selected = True Then
                'Si ce n'est pas une section
245             If lvwSoumission.ListItems(iCompteur).Tag <> vbNullString Then
250               'Si c'est une sous-section
255               If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> vbNullString Then
260                 If m_eType = TYPE_SOUMISSION Then
265                   bSupprimer = True
270                 Else
275                   If lvwSoumission.ListItems(iCompteur).ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_NOIR Or _
                         lvwSoumission.ListItems(iCompteur).ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_BRUN Or _
                         lvwSoumission.ListItems(iCompteur).ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_MAGENTA Then
280                     bSupprimer = True
285                   End If
290                 End If

295                 If bSupprimer = True Then
300                   If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PROFIT) = "" Then
305                     lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PROFIT) = " "
310                   End If

315                   If lvwSoumission.ListItems(iCompteur).ListSubItems(I_COL_SOUM_PROFIT).Tag <> "EXTRA" Then
320                     If m_bModeAjout = False Then
325                       Call AjouterSuppressionCollection(iCompteur)
330                     End If

335                     iIndex = iCompteur

                        'Il faut vérifier si c'est le seul enregistrement de la section.
                        'Si c'est le cas, il faut effacer la section en meme temps
                        'Si l'item sélectionné est le dernier enregistrement
340                     If iIndex = lvwSoumission.ListItems.count Then
                          'Si l'élément d'en haut est une sous-section
345                       If lvwSoumission.ListItems(iIndex - 1).ListSubItems(I_COL_SOUM_PIECE) = vbNullString Then
                            'Il est le seul dans la sous-section
350                         bSeulSS = True
 
                            'Il faut maintenant vérifier si il est le seul dans la section
355                         If lvwSoumission.ListItems(iIndex - 2).Tag = vbNullString Then
                              'Il est le seul enr. dans la section
360                           bSeulS = True
365                         End If
370                       End If
375                     Else
380                       Set itmPrecedent = lvwSoumission.ListItems(iIndex - 1)
385                       Set itmSuivant = lvwSoumission.ListItems(iIndex + 1)

                          'Si l'élément précedent est une sous-section et le suivant est une sous-section ou une section
390                       If itmPrecedent.ListSubItems(I_COL_SOUM_PIECE).Tag = vbNullString And (itmSuivant.Tag = vbNullString Or (itmSuivant.Tag <> vbNullString And itmSuivant.ListSubItems(I_COL_SOUM_PIECE).Tag = vbNullString)) Then
                            'C'est le seul dans la sous-section
395                         bSeulSS = True

                            'Si les éléments avant et après sont des sections
400                         If lvwSoumission.ListItems(iIndex - 2).Tag = vbNullString And itmSuivant.Tag = vbNullString Then
405                           bSeulS = True
410                         End If
415                       End If
420                     End If

425                     Call lvwSoumission.ListItems.Remove(iIndex)

                        'On calcule les prix
430                     Call CalculerPrix

                        'Si c'est le seul dans la sous-section, on efface la sous-section
435                     If bSeulSS = True Then
440                       Call lvwSoumission.ListItems.Remove(iIndex - 1)

445                       iCompteur = iCompteur - 1
450                     End If

                        'Si c'est le seul dans la section, on efface la section
455                     If bSeulS = True Then
460                       Call lvwSoumission.ListItems.Remove(iIndex - 2)

465                       iCompteur = iCompteur - 1
470                     End If
475                   Else
480                     Call MsgBox("La pièce " & lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) & " doit être effacée dans le projet " & lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PROVENANCE), vbOKOnly, "Erreur")

485                     iCompteur = iCompteur + 1
490                   End If
495                 Else
500                   Call MsgBox("La pièce " & lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) & " ne peut pas être supprimée!", vbOKOnly, "Erreur")

505                   iCompteur = iCompteur + 1
510                 End If
515               Else
520                 iCompteur = iCompteur + 1
525               End If
530             Else
535               iCompteur = iCompteur + 1
540             End If
545           Else
550             iCompteur = iCompteur + 1
555           End If
560         Loop
565       Else
            'Cette ligne sert seulement à ne pas déselectionner et repositionner à la ligne 1 si l'utilisateur
            'décide de ne pas supprimer.
            'Le nom de la variable n'est pas significatif dans ce cas, mais c'est celle-ci qui est utilisé pour
            'désélectionner et remettre à la ligne 1
570         bPermission = False
575       End If
580     End If

        'Il faut resélectionner le premier à la fin
585     If lvwSoumission.ListItems.count > 0 Then
590       If bPermission = True Then
595         Call Deselect

600         lvwSoumission.ListItems(1).Selected = True
605       End If
610     End If

615     Exit Sub

AfficherErreur:

620     woups "frmProjSoumMec", "EffacerItemListViewSoumission", Err, Erl
End Sub

Private Sub lvwSoumission_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

5       On Error GoTo AfficherErreur

10      Dim iNbreSelected  As Integer
15      Dim iIndexSelected As Integer
20      Dim iCompteur      As Integer
25      Dim bAfficherMenu  As Boolean

30      If m_eMode = MODE_AJOUT_MODIF Then
35        If Button = vbRightButton Then
40          If lvwSoumission.ListItems.count > 0 Then
              'S'il y a plusieurs items de sélectionnés, c'est parce que l'utilisateur
              'a sélectionné plusieurs items
              'Donc, on ne désélectionne pas
45            For iCompteur = 1 To lvwSoumission.ListItems.count
50              If lvwSoumission.ListItems(iCompteur).Selected = True Then
55                iNbreSelected = iNbreSelected + 1

60                iIndexSelected = iCompteur
65              End If
70            Next

75            If iNbreSelected = 1 Then
80              lvwSoumission.ListItems(iIndexSelected).Selected = False
85            End If

90            Set lvwSoumission.DropHighlight = lvwSoumission.HitTest(X, Y)

95            If Not lvwSoumission.DropHighlight Is Nothing Then
100             If iNbreSelected = 1 Then
105               lvwSoumission.DropHighlight.Selected = True

110               If lvwSoumission.SelectedItem.Tag <> "" Then
115                 If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PROFIT).ForeColor <> COLOR_BLEU And lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PROFIT).ForeColor <> COLOR_ROSE Then
120                   bAfficherMenu = True
125                 Else
130                   bAfficherMenu = False
135                 End If
140               End If
145             Else
150               If m_eType = TYPE_PROJET Then
155                 If lvwSoumission.DropHighlight.Selected = True Then
160                   If lvwSoumission.SelectedItem.Tag <> "" Then
165                     If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PROFIT).Tag <> "EXTRA" Then
170                       If g_bModificationFacturation = True Then
175                         bAfficherMenu = True
180                       Else
185                         bAfficherMenu = False
190                       End If
195                     Else
200                       bAfficherMenu = False
205                     End If
210                   Else
215                     bAfficherMenu = False
220                   End If
225                 End If
230               Else
235                 bAfficherMenu = False
240               End If
245             End If
250           Else
255             bAfficherMenu = False
260           End If

265           If bAfficherMenu = True Then
270             Call RemplirOptionsMenuRightClick(iNbreSelected)

275             Call PopupMenu(mnuRightClick)
280           End If
285         End If
290       Else
295         If Shift <> vbCtrlMask And Shift <> vbShiftMask Then
300           Set lvwSoumission.DropHighlight = Nothing
305         End If
310       End If
315     End If

320     Exit Sub

AfficherErreur:

325     woups "frmProjSoumMec", "lvwSoumission_MouseDown", Err, Erl
End Sub

Private Sub RemplirOptionsMenuRightClick(ByVal iNbreSelected As Integer)

5       On Error GoTo AfficherErreur

10      Dim bFacturer        As Boolean
15      Dim bNC              As Boolean
20      Dim bDateRequise     As Boolean
25      Dim bCommentaire     As Boolean
30      Dim bMauvaisPrix     As Boolean
35      Dim bMaterielInutile As Boolean
40      Dim bTexte           As Boolean
45      Dim bSousSection     As Boolean
50      Dim bFournisseur     As Boolean
55      Dim bAnnulerCommande As Boolean
60      Dim bSupprimer       As Boolean
65      Dim bAjouterPrix     As Boolean
70      Dim bSortieMagasin   As Boolean
75      Dim bQuantite        As Boolean

80      If iNbreSelected > 1 Then
85        If m_eType = TYPE_PROJET Then
90          If g_bModificationFacturation = True Then
95            bFacturer = True
100           bNC = True
105           bSupprimer = True
110         End If
115       End If
120     Else
125       If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) = "" Then
130         bSousSection = True
135       Else
140         If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) = "Texte" Or lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) = "Text" Then
145           bTexte = True

150           If (m_eType = TYPE_SOUMISSION) Or ((m_eType = TYPE_PROJET) And (Right$(txtNoProjSoum.Text, 2) > 19)) Then
155             bSupprimer = True
160           End If
165         Else
170           If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor = -2147483640 Then
175             lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor = 0
180           End If

185           Select Case lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor
                Case COLOR_ORANGE:
190               If g_bModificationFacturation = True Then
195                 bFacturer = True
200                 bNC = True
205               End If

210               bDateRequise = True
215               bCommentaire = True
220               bAnnulerCommande = True
225               bMauvaisPrix = True
                
                Case COLOR_BRUN:
230               bCommentaire = True
235               bFournisseur = True

240               If (m_eType = TYPE_SOUMISSION) Or ((m_eType = TYPE_PROJET) And (Right$(txtNoProjSoum.Text, 2) > 19)) Then
245                 bSupprimer = True
250               End If

                Case COLOR_GRIS:
255               If g_bModificationFacturation = True Then
260                 bFacturer = True
265                 bNC = True
270               End If

275               bCommentaire = True
280               bMauvaisPrix = True
285               bMaterielInutile = True

                Case COLOR_VERT_FORET:
290               bCommentaire = True

295               If (m_eType = TYPE_SOUMISSION) Or ((m_eType = TYPE_PROJET) And (Right$(txtNoProjSoum.Text, 2) > 19)) Then
300                 bSupprimer = True
305               End If

                Case COLOR_ROUGE:
310               bCommentaire = True

                Case COLOR_MAGENTA:
315               bCommentaire = True
320               bAjouterPrix = True

325               If m_eType = TYPE_SOUMISSION Then
330                 bQuantite = True
335               End If

                Case COLOR_NOIR:
340               If m_eType = TYPE_PROJET Then
345                 If g_bModificationFacturation = True Then
350                   bFacturer = True
355                   bNC = True
360                 End If

365                 bMaterielInutile = True
370                 bSortieMagasin = True
375               Else
380                 bQuantite = True
385               End If

400               bCommentaire = True
405               bMauvaisPrix = True
410               bFournisseur = True
415               bSupprimer = True
420           End Select
425         End If
430       End If
435     End If

        'Pour empeche que tous les éléments deviennent invisible, je les mets visible au
        'début
440     mnuFacturer.Visible = True
445     mnuNC.Visible = True
450     mnuDateRequise.Visible = True
455     mnuCommentaire.Visible = True
460     mnuMauvaisPrix.Visible = True
465     mnuInutile.Visible = True
470     mnuTexte.Visible = True
475     mnuChangerSS.Visible = True
480     mnuFournisseur.Visible = True
485     mnuAnnulerCommande.Visible = True
490     mnuSupprimer.Visible = True
495     mnuAjouterPrix.Visible = True
500     mnuSortieMagasin.Visible = True
505     mnuQuantite.Visible = True

510     mnuFacturer.Visible = bFacturer
515     mnuNC.Visible = bNC
520     mnuDateRequise.Visible = bDateRequise
525     mnuCommentaire.Visible = bCommentaire
530     mnuMauvaisPrix.Visible = bMauvaisPrix
535     mnuInutile.Visible = bMaterielInutile
540     mnuTexte.Visible = bTexte
545     mnuChangerSS.Visible = bSousSection
550     mnuFournisseur.Visible = bFournisseur
555     mnuAnnulerCommande.Visible = bAnnulerCommande
560     mnuSupprimer.Visible = bSupprimer
565     mnuAjouterPrix.Visible = bAjouterPrix
570     mnuSortieMagasin.Visible = bSortieMagasin
575     mnuQuantite.Visible = bQuantite

580     Exit Sub

AfficherErreur:

585     woups "frmProjSoumMec", "RemplirOptionsMenuRightClick", Err, Erl
End Sub

Private Sub mnuAjouterPrix_Click()

5       On Error GoTo AfficherErreur

10      Call AjouterPrix

15      Call EnleverSelection

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumMec", "mnuAjouterPrix_Click", Err, Erl
End Sub

Private Sub mnuAnnulerCommande_Click()

5       On Error GoTo AfficherErreur

10      Call AnnulerCommande

15      Call EnleverSelection

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumMec", "mnuAnnulerCommande_Click", Err, Erl
End Sub

Private Sub mnuChangerSS_Click()

5       On Error GoTo AfficherErreur

10      Call ModifierSousSection

15      Call EnleverSelection

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumMec", "mnuChangerSS_Click", Err, Erl
End Sub

Private Sub mnuDateRequise_Click()

5       On Error GoTo AfficherErreur

10      If Trim$(lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_DATE_REQUISE)) = "" Then
15        mvwDateRequise.Year = Year(Date)
20        mvwDateRequise.Month = Month(Date)
25        mvwDateRequise.Day = Day(Date)
30      Else
35        mvwDateRequise.Year = Left$(lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_DATE_REQUISE), 4)
40        mvwDateRequise.Month = Mid$(lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_DATE_REQUISE), 6, 2)
45        mvwDateRequise.Day = Right$(lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_DATE_REQUISE), 2)
50      End If

55      fraDateRequise.Top = lvwSoumission.Top

60      fraDateRequise.Visible = True

65      Exit Sub

AfficherErreur:

70      woups "frmProjSoumMec", "mnuDateRequise_Click", Err, Erl
End Sub

Private Sub mnuCommentaire_Click()

5       On Error GoTo AfficherErreur

10      txtcommentaire.Text = lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_COMMENTAIRE)

15      fraCommentaire.Top = lvwSoumission.Top

20      fraCommentaire.Visible = True

25      Exit Sub

AfficherErreur:

30      woups "frmProjSoumMec", "mnuCommentaire_Click", Err, Erl
End Sub

Private Sub mnuFacturer_Click()

5       On Error GoTo AfficherErreur

10      Call FacturerDate

15      Call EnleverSelection

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumMec", "mnuFacturer_Click", Err, Erl
End Sub

Private Sub mnuFournisseur_Click()

5       On Error GoTo AfficherErreur

10      Call ChangerFournisseurRetour

15      Call EnleverSelection

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumMec", "mnuFournisseur_Click", Err, Erl
End Sub

Private Sub mnuInutile_Click()

5       On Error GoTo AfficherErreur

10      Call MaterielInutile

15      Call EnleverSelection

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumMec", "mnuInutile_Click", Err, Erl
End Sub

Private Sub mnuMauvaisPrix_Click()

5       On Error GoTo AfficherErreur

10      Call MauvaisPrix

15      Call EnleverSelection

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumMec", "mnuMauvaisPrix", Err, Erl
End Sub

Private Sub mnuNC_Click()

5       On Error GoTo AfficherErreur

10      Call FacturerNC

15      Call EnleverSelection

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumMec", "mnuNC_Click", Err, Erl
End Sub

Private Sub mnuQuantite_Click()

5       On Error GoTo AfficherErreur

10      Call ChangerQuantite

15      Call EnleverSelection

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumMec", "mnuQuantite_Click", Err, Erl
End Sub

Private Sub mnuSortieMagasin_Click()

5       On Error GoTo AfficherErreur

10      Call SortieMagasin

15      Call EnleverSelection

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumMec", "mnuSortieMagasin_Click", Err, Erl
End Sub

Private Sub mnuSupprimer_Click()
  
5       On Error GoTo AfficherErreur

10      Call EffacerItemListViewSoumission

15      Call EnleverSelection

20      Exit Sub
  
AfficherErreur:

25      woups "frmProjSoumMec", "mnuSupprimer_Click", Err, Erl
End Sub

Private Sub mnuTexte_Click()

5       On Error GoTo AfficherErreur

10      Call ModifierTexte

15      Call EnleverSelection

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumMec", "mnuTexte_Click", Err, Erl
End Sub

Private Sub EnleverSelection()

5       On Error GoTo AfficherErreur

10      Set lvwSoumission.DropHighlight = Nothing

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMec", "EnleverSelection", Err, Erl
End Sub

Private Sub mvwDateRequise_GotFocus()
        
5       On Error GoTo AfficherErreur

10      m_bMonthViewHasFocus = True

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMec", "mvwDateRequise_GotFocus", Err, Erl
End Sub

Private Sub txtCheminPhotos_KeyDown(KeyCode As Integer, Shift As Integer)

5       On Error GoTo AfficherErreur

10      If m_eMode = MODE_AJOUT_MODIF Then
15        If KeyCode <> vbKeyBack And KeyCode <> vbKeyDelete Then
20          KeyCode = 0
25        End If
30      End If

35      Exit Sub

AfficherErreur:

40      woups "frmProjSoumMec", "txtCheminPhotos_KeyDown", Err, Erl
End Sub



Private Sub txtPrixManuel_LostFocus()
  
5       On Error GoTo AfficherErreur
  
10      txtPrixManuel.Text = Replace(txtPrixManuel.Text, ".", ",")
  
15      Call CalculerPrix

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumMec", "txtPrixManuel_LostFocus", Err, Erl
End Sub

Private Sub txtPrixList_LostFocus()

5       On Error GoTo AfficherErreur

10      If txtPrixList.Text <> vbNullString Then
15        txtPrixList.Text = Replace(txtPrixList, ".", ",")
  
20        If IsNumeric(txtPrixList.Text) Then
25          Call CalculerPrixNet
30        Else
35          Call MsgBox("Valeur non numérique!", vbOKOnly, "Erreur")
40          txtPrixList.Text = vbNullString
45        End If
50      End If

55      Exit Sub

AfficherErreur:

60     woups "frmProjSoumMec", "txtPrixList_LostFocus", Err, Erl
End Sub

Private Sub txtPrixNet_Change()

5       On Error GoTo AfficherErreur
 
        'Quand le contenu du prix net change
  
        'Si la longueur du texte écrit est plus grand que 0
10      If Len(txtPrixNet.Text) > 0 Then
          'On vide le prix spécial et on le désactive
15        txtPrixSpecial.Text = vbNullString
20        txtPrixSpecial.Enabled = False
25      Else
          'Sinon, on active le prix spécial
30        txtPrixSpecial.Enabled = True
35      End If

40      Exit Sub

AfficherErreur:

45      woups "frmProjSoumMec", "txtPrixNet_Change", Err, Erl
End Sub

Private Sub txtPrixNet_GotFocus()

5       On Error GoTo AfficherErreur

        'Si le prix net prend le focus
10      Call CalculerPrixNet

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMec", "txtPrixNet_GotFocus", Err, Erl
End Sub

Private Sub CalculerPrixNet()

5       On Error GoTo AfficherErreur

10      Dim dblEscompte As Double
15      Dim dblPrix     As Double
  
        'Si le prix net n'est pas barré.. ie.. si le prix spécial est vide
20      If txtPrixNet.Locked = False Then
25        mskEscompte.Text = Replace(mskEscompte.Text, "_", vbNullString)
    
30        mskEscompte.Text = Replace(mskEscompte.Text, ".", ",")
    
35        If mskEscompte.Text <> vbNullString Then
40          dblEscompte = CDbl(mskEscompte.Text)
45        Else
50          dblEscompte = 0
55        End If
              
60        If txtPrixList.Text <> vbNullString Then
65          dblPrix = CDbl(Replace(txtPrixList.Text, ".", ","))
70        Else
75          dblPrix = 0
80        End If
    
          'Calcul du prix net
85        txtPrixNet.Text = Round((dblPrix) * (1 - dblEscompte), 4)
    
90        txtPrixNet.Text = Replace(txtPrixNet.Text, ".", ",")
95      End If

100     Exit Sub

AfficherErreur:

105     woups "frmProjSoumMec", "CalculerPrixNet", Err, Erl
End Sub

Private Sub txtPrixNet_LostFocus()

5       On Error GoTo AfficherErreur

10      txtPrixNet.Text = Replace(txtPrixNet.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMec", "txtPrixNet_LostFocus", Err, Erl
End Sub

Private Sub ViderChamps_frs()

5       On Error GoTo AfficherErreur

        'Vide les champs pieces
10      txtPrixList.Text = vbNullString
15      mskEscompte.Text = vbNullString
20      txtPrixNet.Text = vbNullString
25      txtPrixSpecial.Text = vbNullString
  
30      optCAN.Value = True

35      Call AfficherDrapeau

40      Exit Sub

AfficherErreur:

45      woups "frmProjSoumMec", "ViderChamps_frs", Err, Erl
End Sub

Private Sub ModifierPrixCatalogue()
        'Enregistrement du prix de la pièce
  
5       On Error GoTo AfficherErreur

10      Dim rstPrix     As ADODB.Recordset
15      Dim dblPrixList As Double
20      Dim dblEscompte As Double
25      Dim dblPrixNet  As Double
        
30      If txtPrixList.Text <> "" Then
35        dblPrixList = CDbl(txtPrixList.Text)
40      Else
45        dblPrixList = 0
50      End If
        
55      If mskEscompte.Text <> vbNullString Then
60        dblEscompte = CDbl(mskEscompte.Text)
65      Else
70        dblEscompte = 0
75      End If
        
80      If txtPrixNet.Text <> "" Then
85        dblPrixNet = CDbl(txtPrixNet.Text)
90      Else
95        dblPrixNet = CDbl(txtPrixSpecial.Text)
100     End If

105     Set rstPrix = New ADODB.Recordset
        
110     If txtPrixNet.Enabled = True Then
          'Ouverture du recordset
115       Call rstPrix.Open("SELECT * FROM GRB_PiecesFRS WHERE PIECE = '" & Replace(lvwSoumission.ListItems(CInt(fraPrix.Tag)).SubItems(I_COL_SOUM_PIECE), "'", "''") & "' AND IDFRS = " & cmbfrs.ItemData(cmbfrs.ListIndex) & " AND PRIX_NET <> ''", g_connData, adOpenDynamic, adLockOptimistic)

120       If rstPrix.EOF Then
125         Call rstPrix.AddNew

130         rstPrix.Fields("PIECE").Value = lvwSoumission.ListItems(CInt(fraPrix.Tag)).SubItems(I_COL_SOUM_PIECE)
135         rstPrix.Fields("IDFRS").Value = cmbfrs.ItemData(cmbfrs.ListIndex)
140       End If

145       rstPrix.Fields("PRIX_LIST") = dblPrixList
150       rstPrix.Fields("ESCOMPTE") = dblEscompte
155       rstPrix.Fields("PRIX_NET") = dblPrixNet
160       rstPrix.Fields("PRIX_SP") = ""
165     Else
170       Call rstPrix.Open("SELECT * FROM GRB_PiecesFRS WHERE PIECE = '" & Replace(lvwSoumission.ListItems(CInt(fraPrix.Tag)).SubItems(I_COL_SOUM_PIECE), "'", "''") & "' AND IDFRS = " & cmbfrs.ItemData(cmbfrs.ListIndex) & " AND PRIX_SP <> ''", g_connData, adOpenDynamic, adLockOptimistic)

175       If rstPrix.EOF Then
180         Call rstPrix.AddNew

185         rstPrix.Fields("PIECE").Value = lvwSoumission.ListItems(CInt(fraPrix.Tag)).SubItems(I_COL_SOUM_PIECE)
190         rstPrix.Fields("IDFRS").Value = cmbfrs.ItemData(cmbfrs.ListIndex)
195       End If

200       rstPrix.Fields("PRIX_SP") = dblPrixNet
205       rstPrix.Fields("PRIX_LIST") = ""
210       rstPrix.Fields("ESCOMPTE") = ""
215       rstPrix.Fields("PRIX_NET") = ""
220     End If

225     If optCAN.Value = True Then
230       rstPrix.Fields("DeviseMonétaire") = "CAN"
235     Else
240       If optUSA.Value = True Then
245         rstPrix.Fields("DeviseMonétaire") = "USA"
250       Else
255         rstPrix.Fields("DeviseMonétaire") = "SPA"
260       End If
265     End If

270     rstPrix.Fields("Type") = "M"

275     rstPrix.Fields("ENTRER_PAR") = g_sInitiale

280     rstPrix.Fields("Date") = ConvertDate(Date)
  
285     Call rstPrix.Update
  
290     Call rstPrix.Close
295     Set rstPrix = Nothing
        
300     Exit Sub

AfficherErreur:

305     woups "frmProjSoumMec", "ModifierPrixCatalogue", Err, Erl
End Sub

Private Sub optCAN_Click()

5       On Error GoTo AfficherErreur

        'dependant la devise, affiche le drapeau
10      Call AfficherDrapeau

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMec", "optCAN_Click", Err, Erl
End Sub
            
Private Sub AfficherDrapeau()

5       On Error GoTo AfficherErreur

        '''''''''''''''''''''''''''''''''''''''''
        'Dépendant la devise, affiche le drapeau'
        '''''''''''''''''''''''''''''''''''''''''
10      If optCAN.Value = True Then
15        imgCanada.Visible = True
20        imgEU.Visible = False
25        imgSpain.Visible = False
30      Else
35        If optUSA.Value = True Then
40          imgEU.Visible = True
45          imgCanada.Visible = False
50          imgSpain.Visible = False
55        Else
60          imgSpain.Visible = True
65          imgCanada.Visible = False
70          imgEU.Visible = False
75        End If
80      End If

85      Exit Sub

AfficherErreur:

90     woups "frmProjSoumMec", "AfficherDrapeau", Err, Erl
End Sub

Private Sub optSpain_Click()

5       On Error GoTo AfficherErreur

        'dependant la devise, affiche le drapeau
10      Call AfficherDrapeau

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMec", "optSpain_Click", Err, Erl
End Sub

Private Sub optUSA_Click()

5       On Error GoTo AfficherErreur

        'dependant la devise, affiche le drapeau
10      Call AfficherDrapeau

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMec", "optUSA_Click", Err, Erl
End Sub

Private Sub mskEscompte_GotFocus()

5       On Error GoTo AfficherErreur

        'Quand le maskEdit prend le focus, on set le masque
10      If mskEscompte.Enabled = True Then
15        mskEscompte.mask = "0,####"
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmProjSoumMec", "mskEscompte_GotFocus", Err, Erl
End Sub

Private Sub mskEscompte_LostFocus()

5      On Error GoTo AfficherErreur

        'Quand le maskEdit perd le focus, on enlève le mask
10      mskEscompte.mask = vbNullString
  
        'Si le champs contient 0,____, c'est parce que rien n'a été entré
15      If mskEscompte.Text = "0,____" Then
          'Donc, on le vide
20        mskEscompte.Text = vbNullString
25      End If
  
30      Call CalculerPrixNet

35      Exit Sub

AfficherErreur:

40      woups "frmProjSoumMec", "mskEscompte_LostFocus", Err, Erl
End Sub

Private Sub cmdExtra_Click()

5       On Error GoTo AfficherErreur

10      Dim rstProjSoum As ADODB.Recordset
15      Dim sNumero     As String
20      Dim bExiste     As Boolean
25      Dim sExtension  As String
30      Dim bNoValide   As Boolean

35      If Right$(txtNoProjSoum.Text, 2) = "99" Then
40        Call MsgBox("Vous ne pouvez pas faire un extra à partir de ce projet!", vbOKOnly, "Erreur")

45        Exit Sub
50      End If

55      sExtension = Right$(txtNoProjSoum.Text, 2)
        
60      sNumero = InputBox("Quel est l'extension à ajouter au numéro " & Left$(txtNoProjSoum.Text, 6) & "?")
  
65      If sNumero <> vbNullString Then
70        If Not IsNumeric(sNumero) Then
75          Call MsgBox("Numéro non numérique!", vbOKOnly, "Erreur")
        
80          Exit Sub
85        End If
    
90        sNumero = Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 3) & "-" & sNumero
    
95        Screen.MousePointer = vbHourglass

100       bNoValide = True

105       If ValiderFormatNumeroProjSoum(sNumero) = False Then
110         bNoValide = False
115       End If

120       If bNoValide = True Then
125         If ValiderFormatMecanique(sNumero) = False Then
130           bNoValide = False
135         End If
140       End If

145       If bNoValide = True Then
150         If ValiderFormatJobExtra(sNumero) = False Then
155           bNoValide = False
160         End If
165       End If

170       If bNoValide = False Then
175         Screen.MousePointer = vbDefault

180         Exit Sub
185       End If

190       sNumero = UCase(sNumero)
  
195       Set rstProjSoum = New ADODB.Recordset

200       Call rstProjSoum.Open("SELECT * FROM GRB_SoumissionElec WHERE IDSoumission = '" & sNumero & "'", g_connData, adOpenDynamic, adLockOptimistic)
        
205       If rstProjSoum.EOF Then
210         bExiste = False
215       Else
220          bExiste = True

225         Call MsgBox("Ce numéro existe dans les soumissions électriques!", vbOKOnly, "Erreur")
230       End If

235       Call rstProjSoum.Close

240       If bExiste = False Then
245         Call rstProjSoum.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & sNumero & "'", g_connData, adOpenDynamic, adLockOptimistic)
        
250         If rstProjSoum.EOF Then
255           bExiste = False
260         Else
265           bExiste = True

270           Call MsgBox("Ce numéro existe dans les projets électriques!", vbOKOnly, "Erreur")
275         End If

280         Call rstProjSoum.Close
285       End If

290       If bExiste = False Then
295         Call rstProjSoum.Open("SELECT * FROM GRB_SoumissionMec WHERE IDSoumission = '" & sNumero & "'", g_connData, adOpenDynamic, adLockOptimistic)

300         If rstProjSoum.EOF Then
305           bExiste = False
310         Else
315           bExiste = True

320           Call MsgBox("Ce numéro existe dans les soumissions mécaniques!", vbOKOnly, "Erreur")
325         End If

330         Call rstProjSoum.Close
335       End If
          
340       If bExiste = False Then
345         Call rstProjSoum.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = '" & sNumero & "'", g_connData, adOpenDynamic, adLockOptimistic)

350         If rstProjSoum.EOF Then
355           bExiste = False
360         Else
365           bExiste = True

370           Call MsgBox("Ce numéro existe dans les projets mécaniques!", vbOKOnly, "Erreur")
375         End If

380         Call rstProjSoum.Close
385       End If
    
          'Si le projet ou la soumission n'existe pas
390       If bExiste = False Then
            'Pour pouvoir réafficher l'élément affiché avant l'ajout au cas où la personne
            'annule l'ajout
395         Call rstProjSoum.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & sNumero & "'", g_connData, adOpenDynamic, adLockOptimistic)

400         If Not rstProjSoum.EOF Then
405           If rstProjSoum.Fields("Ouvert") = False Then
410             Call MsgBox("Ce numéro est fermé!", vbOKOnly, "Erreur")

415             Call rstProjSoum.Close
420             Set rstProjSoum = Nothing

425             Screen.MousePointer = vbDefault

430             Exit Sub
435           End If
440         End If

445         Call rstProjSoum.Close
450         Set rstProjSoum = Nothing
              
455         m_sAncienProjSoum = txtNoProjSoum.Text

460         Call InitialiserVariables(txtNoProjSoum.Text)
      
            'Affiche le nouveau numéro
465         txtNoProjSoum.Text = sNumero

470         If Right$(sNumero, 2) >= 60 And Right$(sNumero, 2) <= 98 Then
475           m_sLiaison = InputBox("Quelle est l'extention au projet " & Left$(txtNoProjSoum.Text, 6) & " auquel ce projet sera lié?", , sExtension)
480         End If

485         m_bModeAjout = True
  
            'Vide le listview
490         Call lvwSoumission.ListItems.Clear

495         Call CalculerPrix
  
            'Débarre les champs
500         Call BarrerChamps(False)
      
505         m_sTempsDessin = 0
510         m_sTempsCoupe = 0
515         m_sTempsMachinage = 0
520         m_sTempsSoudure = 0
525         m_sTempsAssemblage = 0
530         m_sTempsPeinture = 0
535         m_sTempsTest = 0
540         m_sTempsInstallation = 0
545         m_sTempsFormation = 0
550         m_sTempsGestion = 0

555         m_sNbrePersonne = 0
560         m_sTempsHebergement = 0
565         m_sTempsRepas = 0
570         m_sTempsTransport = 0
575         m_sTempsUniteMobile = 0
580         m_sPrixEmballage = 0
      
585         txtNbreManuel.Text = 0
590         txtPrixManuel.Text = 0

595         txtForfait.Text = ""
600         lblForfaitInitiale.Caption = ""

605         txtPrixReception.Text = "0"
610         txtPrixSoumission.Text = "0"
    
615         txtPrixTotal.Text = "0"
620         txtProfit.Text = "0"
625         txtCommission.Text = "0"
630         txtTotalTemps.Text = "0"
635         txtTotalPieces.Text = "0"
640         txtImprevus.Text = "0"
645         txtNoSoumission.Text = vbNullString
      
            'Vide la valeur par défaut si demande Sous-Section
650         m_sSousSection = vbNullString

655         txtDescription.Text = vbNullString

660         m_bModeAjout = True
665         m_bModeAffichage = False
670         m_bExtra = True

675         lvwSoumission.Height = lvwSoumission.Height * 0.49
680         lvwSoumission.Top = lvwPieces.Top + lvwPieces.Height + (lvwSoumission.Height * 0.02)
        
            'Met le form en mode ajout/modif
685         Call AfficherControles(MODE_AJOUT_MODIF)
690       End If
  
695       Screen.MousePointer = vbDefault
700     End If

705     Exit Sub

AfficherErreur:

710     woups "frmProjSoumMec", "cmdAjouter_Click", Err, Erl
End Sub

Private Function VerifierSiOuvert(ByRef sUser As String) As Boolean
        'Vérifie si le projet ou la soumission n'est pas en modification
        'par un autre utilisateur sur un autre ordinateur
5       On Error GoTo AfficherErreur

10      Dim rstProjSoum   As ADODB.Recordset
15      Dim bModification As Boolean

20      Set rstProjSoum = New ADODB.Recordset

25      If m_eType = TYPE_PROJET Then
30        Call rstProjSoum.Open("SELECT Modification, Par FROM GRB_ProjetMec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
35      Else
40        Call rstProjSoum.Open("SELECT Modification, Par FROM GRB_SoumissionMec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
45      End If

50      If rstProjSoum.Fields("Modification") = True Then
55        sUser = rstProjSoum.Fields("Par")
60        bModification = True
65      Else
70        sUser = ""
75        bModification = False
80      End If

85      Call rstProjSoum.Close
90      Set rstProjSoum = Nothing

95      VerifierSiOuvert = bModification

100     Exit Function

AfficherErreur:

105     woups "frmProjSoumMec", "VerifierSiOuvert", Err, Erl
End Function

Private Sub OuvrirProjSoum(ByVal bOuvrir As Boolean)
        'Remplit ou vide les champs Modification et Par
5       On Error GoTo AfficherErreur

10      Dim rstProjSoum As ADODB.Recordset

15      Set rstProjSoum = New ADODB.Recordset

20      If IsNumeric(Right$(txtNoProjSoum.Text, 2)) Then
25        If ((Right$(txtNoProjSoum.Text, 2) >= "00" And Right$(txtNoProjSoum.Text, 2) <= "19") Or (Right$(txtNoProjSoum.Text, 2) >= "80" And Right$(txtNoProjSoum.Text, 2) <= "98")) And (m_eType = TYPE_PROJET) Then
30          If Right$(txtNoProjSoum.Text, 2) >= "00" And Right$(txtNoProjSoum.Text, 2) <= "19" Then
35            Call rstProjSoum.Open("SELECT Modification, Par FROM GRB_ProjetMec WHERE IDProjet = '" & txtNoProjSoum.Text & "' OR IDProjet = '" & Left$(txtNoProjSoum.Text, 6) & "-" & Right$(txtNoProjSoum.Text, 2) + 80 & "'", g_connData, adOpenDynamic, adLockOptimistic)
40          Else
45            Call rstProjSoum.Open("SELECT Modification, Par FROM GRB_ProjetMec WHERE IDProjet = '" & txtNoProjSoum.Text & "' OR IDProjet = '" & Left$(txtNoProjSoum.Text, 6) & "-" & Right$("0" & Right$(txtNoProjSoum.Text, 2) - 80, 2) & "'", g_connData, adOpenDynamic, adLockOptimistic)
50          End If
55        Else
60          If m_eType = TYPE_PROJET Then
65            Call rstProjSoum.Open("SELECT Modification, Par FROM GRB_ProjetMec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
70          Else
75            Call rstProjSoum.Open("SELECT Modification, Par FROM GRB_SoumissionMec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
80          End If
85        End If
90      Else
95        If m_eType = TYPE_PROJET Then
100         Call rstProjSoum.Open("SELECT Modification, Par FROM GRB_ProjetMec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
105       Else
110         Call rstProjSoum.Open("SELECT Modification, Par FROM GRB_SoumissionMec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
115       End If
120     End If

125     Do While Not rstProjSoum.EOF
130       If bOuvrir = True Then
135         rstProjSoum.Fields("Modification") = True
140         rstProjSoum.Fields("Par") = g_sEmploye
145       Else
150         rstProjSoum.Fields("Modification") = False
155         rstProjSoum.Fields("Par") = ""
160       End If

165       Call rstProjSoum.Update

170       Call rstProjSoum.MoveNext
175     Loop

180     Call rstProjSoum.Close
185     Set rstProjSoum = Nothing

190     Exit Sub

AfficherErreur:

195     woups "frmProjSoumMec", "OuvrirProjSoum", Err, Erl
End Sub

Private Sub cmdReset_Click()
        'Permet d'effacer le champs Modification et Par si c'est le user actuel
5       On Error GoTo AfficherErreur

10      Dim rstProjSoum As ADODB.Recordset

15      Set rstProjSoum = New ADODB.Recordset

20      If MsgBox("Êtes-vous certains de ne pas être en modification sur un autre ordinateur?", vbYesNo) = vbYes Then
25        If m_eType = TYPE_PROJET Then
30          Call rstProjSoum.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
35        Else
40          Call rstProjSoum.Open("SELECT * FROM GRB_SoumissionMec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
45        End If

50        rstProjSoum.Fields("Modification") = False
55        rstProjSoum.Fields("Par") = ""

60        Call rstProjSoum.Update

65        Call rstProjSoum.Close
70        Set rstProjSoum = Nothing

75        cmdReset.Visible = False
80      End If

85      Exit Sub

AfficherErreur:

90      woups "frmProjSoumMec", "cmdReset_Click", Err, Erl
End Sub

Private Sub lvwPieceTrouve_DblClick()

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer

15      m_bRecherchePiece = True
20      m_bPieceInutile = False
25      m_bChangementFRS = False

        'On affiche lvwFournisseur selon la pièce choisie
        'Rempli les fournisseurs de la pièce choisie
30      Call AfficherListeFournisseurs
  
        'si le listview n'est pas vide
35      If lvwfournisseur.ListItems.count = 1 Then
40        If MsgBox("Il n'y a aucun fournisseur pour cette pièce!" & vbNewLine & "Voulez-vous en ajouter?", vbYesNo, "Erreur") = vbYes Then
45          Screen.MousePointer = vbHourglass
      
            'On ouvre le catalogue sur cet enregistrement
50          Call FrmCatalogueElec.AfficherForm(lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_CATEGORIE), lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_MANUFACT), lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM))
      
55          Screen.MousePointer = vbDefault
      
            'On rappelle la méthode
60          Call AfficherListeFournisseurs
65        End If
70      End If

75      Exit Sub
 
AfficherErreur:

80      woups "frmProjSoumMec", "lvwPieceTrouve_DblClick", Err, Erl
End Sub

Private Sub cmdOKPieceTrouve_Click()

5       On Error GoTo AfficherErreur

10      m_bRecherchePiece = True
15      m_bPieceInutile = False

        'On affiche lvwFournisseur selon la pièce choisie
        'Rempli les fournisseurs de la pièce choisie
20      Call AfficherListeFournisseurs
  
        'si le listview n'est pas vide
25      If lvwfournisseur.ListItems.count = 1 Then
30        If MsgBox("Il n'y a aucun fournisseur pour cette pièce!" & vbNewLine & "Voulez-vous en ajouter?", vbYesNo, "Erreur") = vbYes Then
35          Screen.MousePointer = vbHourglass
      
            'On ouvre le catalogue sur cet enregistrement
40          Call FrmCatalogueElec.AfficherForm(lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_CATEGORIE), lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_MANUFACT), lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM))
      
45          Screen.MousePointer = vbDefault
      
            'On rappelle la méthode
50          Call AfficherListeFournisseurs
55        End If
60      End If

65      Exit Sub

AfficherErreur:

70      woups "frmProjSoumMec", "cmdOKPieceTrouve_Click", Err, Erl
End Sub

Private Sub cmdAnnulerPieceTrouve_Click()

5       On Error GoTo AfficherErreur

10      fraPieceTrouve.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMec", "cmdAnnulerPieceTrouve", Err, Erl
End Sub

Private Sub RemplirListViewRecherche(ByVal iIndexColumn As Integer, ByVal sTexte As String)

5       On Error GoTo AfficherErreur

10      Dim rstPiece   As ADODB.Recordset
15      Dim itmPiece   As ListItem
20      Dim iCompteur  As Integer
25      Dim sChamps    As String
30      Dim sRecherche As String
35      Dim sLettre    As String

40      Call lvwPieceTrouve.ListItems.Clear

45      If iIndexColumn = I_COL_PIECES_NO_ITEM Then
50        For iCompteur = 1 To Len(sTexte)
55          sLettre = Mid$(sTexte, iCompteur, 1)

60          If (Asc(sLettre) >= 48 And Asc(sLettre) <= 57) Or _
               (Asc(sLettre) >= 65 And Asc(sLettre) <= 90) Or _
               (Asc(sLettre) >= 97 And Asc(sLettre) <= 122) Then
65            sRecherche = sRecherche & sLettre
70          End If
75        Next
80      End If

        'Attribue le nom du champs selon la colonne cliquée
85      Select Case iIndexColumn
          Case I_COL_PIECES_PIECE_GRB: sChamps = "PIECE_GRB"
90        Case I_COL_PIECES_NO_ITEM:   sChamps = "PIECE_MODIF"
95        Case I_COL_PIECES_DESCR_EN:  sChamps = "DESC_EN"
100       Case I_COL_PIECES_DESCR_FR:  sChamps = "DESC_FR"
105       Case I_COL_PIECES_MANUFACT:  sChamps = "FABRICANT"
110     End Select
        
115     Set rstPiece = New ADODB.Recordset
        
120     If iIndexColumn = I_COL_PIECES_NO_ITEM Then
125       Call rstPiece.Open("SELECT * FROM GRB_CatalogueMec WHERE INSTR(1," & sChamps & ",'" & sRecherche & "')> 0 ", g_connData, adOpenDynamic, adLockOptimistic)
130     Else
135       Call rstPiece.Open("SELECT * FROM GRB_CatalogueMec WHERE INSTR(1," & sChamps & ",'" & Replace(sTexte, "'", "''") & "')> 0 ", g_connData, adOpenDynamic, adLockOptimistic)
140     End If
          
        'Pour chaque enregistrement
145     Do While Not rstPiece.EOF
          'On ajoute dans le ListView
150       Set itmPiece = lvwPieceTrouve.ListItems.Add

155       If Not IsNull(rstPiece.Fields("PIECE_GRB")) Then
160         itmPiece.Text = rstPiece.Fields("PIECE_GRB")
165       Else
170         itmPiece.Text = ""
175       End If

180       itmPiece.SubItems(I_COL_RECH_NO_ITEM) = rstPiece.Fields("PIECE")
185       itmPiece.SubItems(I_COL_RECH_CATEGORIE) = rstPiece.Fields("CATEGORIE")

190       If Not IsNull(rstPiece.Fields("FABRICANT")) Then
195         itmPiece.SubItems(I_COL_RECH_MANUFACT) = rstPiece.Fields("FABRICANT")
200       Else
205         itmPiece.SubItems(I_COL_RECH_MANUFACT) = ""
210       End If

215       If Not IsNull(rstPiece.Fields("DESC_EN")) Then
220         itmPiece.SubItems(I_COL_RECH_DESCR_EN) = rstPiece.Fields("DESC_EN")
225       Else
230         itmPiece.SubItems(I_COL_RECH_DESCR_EN) = ""
235       End If

240       If Not IsNull(rstPiece.Fields("DESC_FR")) Then
245         itmPiece.SubItems(I_COL_RECH_DESCR_FR) = rstPiece.Fields("DESC_FR")
250       Else
255         itmPiece.SubItems(I_COL_RECH_DESCR_FR) = ""
260       End If

265       Call rstPiece.MoveNext
270     Loop

275     Call rstPiece.Close
280     Set rstPiece = Nothing

285     Exit Sub

AfficherErreur:

290     woups "frmProjSoumMec", "RemplirListViewRecherche", Err, Erl
End Sub

Private Sub mvwDateFacturation_LostFocus()

5       On Error GoTo AfficherErreur

        'Quand le calendrier perd le focus, il faut l'enlever
10      mvwDateFacturation.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMec", "mvwDateFacturation_LostFocus", Err, Erl
End Sub

Private Sub mvwDateFacturation_DateClick(ByVal DateClicked As Date)

5       On Error GoTo AfficherErreur

        'Affiche la date dans le TextBox sous le format AAAA-MM-JJ
10      txtDateFacturation.Text = ConvertDate(DateClicked)

        'Enlever le calendrier
15      mvwDateFacturation.Visible = False

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumMec", "mvwDateFacturation_DateClick", Err, Erl
End Sub

Private Sub ImprimerProjSoumFacturation(ByVal rstProjSoum As ADODB.Recordset, ByVal sNoFacture As String)

5       On Error GoTo AfficherErreur

        'Impression de la feuille de soumission
10      Dim rstPiece             As ADODB.Recordset
15      Dim rstTemp              As ADODB.Recordset
20      Dim rstImpProjSoum       As ADODB.Recordset
25      Dim rstSoum              As ADODB.Recordset
30      Dim sOrdreSection        As String
35      Dim iCompteurSoum        As Integer
40      Dim sSousSection         As String
45      Dim sSousSectionRS       As String
50      Dim dblTempsDessin       As Double
55      Dim dblTempsCoupe        As Double
60      Dim dblTempsMachinage    As Double
65      Dim dblTempsSoudure      As Double
70      Dim dblTempsAssemblage   As Double
75      Dim dblTempsPeinture     As Double
80      Dim dblTempsTest         As Double
85      Dim dblTempsInstallation As Double
90      Dim dblTempsFormation    As Double
95      Dim dblTempsGestion      As Double
100     Dim dblTempsShipping     As Double
105     Dim dblTotalTemps        As Double
110     Dim sChampsSection       As String
115     Dim sNoProjet            As String
120     Dim sNoSoumission        As String
125     Dim dblPrixEmballage     As Double
130     Dim sTotalTemps          As String
135     Dim sTotalPiece          As String
140     Dim sProfit              As String
145     Dim sImprevue            As String
150     Dim sCommission          As String
155     Dim sAutre               As String
160     Dim sPrixTotal           As String

        'Supprime les données de l'impression
165     Call g_connData.Execute("DELETE * FROM GRB_Impression_Soumission")

170     iCompteurSoum = 1
  
175     Screen.MousePointer = vbHourglass

180     Set rstImpProjSoum = New ADODB.Recordset

185     Call rstImpProjSoum.Open("SELECT * FROM GRB_impression_soumission", g_connData, adOpenDynamic, adLockOptimistic)
    
190     sOrdreSection = vbNullString
            
195     sNoProjet = rstProjSoum.Fields("IDProjet")
200     sNoSoumission = rstProjSoum.Fields("IDSoumission")

205     Set rstPiece = New ADODB.Recordset

210     Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoProjet & "' And Type = 'M' And Facturation = '" & sNoFacture & "'", g_connData, adOpenDynamic, adLockOptimistic)

215     Set rstTemp = New ADODB.Recordset
  
220     Do While Not rstPiece.EOF
225       sSousSectionRS = rstPiece.Fields("SousSection")
       
230       If sSousSectionRS = S_PAS_SOUS_SECTION Then
235         sSousSectionRS = " "
240       End If
      
245       If sOrdreSection <> rstPiece.Fields("OrdreSection") Then
            'remplis la table impression_soumission
            'ajoute seulement la section
250         sOrdreSection = rstPiece.Fields("OrdreSection")

255         If m_eLangage = ANGLAIS Then
260           sChampsSection = "NomSectionEN"
265         Else
270           sChampsSection = "NomSectionFR"
275         End If

280         Call rstTemp.Open("SELECT " & sChampsSection & " FROM GRB_SoumProjSectionMec WHERE IDSection = " & rstPiece.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
          
            'Ajoute la section dans la soumission
285         Call rstImpProjSoum.AddNew
          
290         rstImpProjSoum.Fields("NoLigne") = iCompteurSoum

295         rstImpProjSoum.Fields("IDSoumission") = sNoProjet
        
300         If Not IsNull(rstTemp.Fields(sChampsSection)) Then
305           rstImpProjSoum.Fields("NomSection") = rstTemp.Fields(sChampsSection)
310         Else
315           rstImpProjSoum.Fields("NomSection") = " "
320         End If
           
325         Call rstImpProjSoum.Update
         
330         iCompteurSoum = iCompteurSoum + 1
       
335         Call rstTemp.Close
          
340         sSousSection = rstPiece.Fields("SousSection")
         
345         If sSousSection = S_PAS_SOUS_SECTION Then
350           sSousSection = " "
355         End If
          
360         Call rstImpProjSoum.AddNew

365         rstImpProjSoum.Fields("NoLigne") = iCompteurSoum

370         rstImpProjSoum.Fields("IDSoumission") = sNoProjet

375         rstImpProjSoum.Fields("SousSection") = sSousSection
        
380         Call rstImpProjSoum.Update
          
385         iCompteurSoum = iCompteurSoum + 1
390       Else
            'ajoute une soussection dans impression_soum
395         If sSousSection <> sSousSectionRS Then
400           sSousSection = sSousSectionRS
       
405           Call rstImpProjSoum.AddNew
      
410           rstImpProjSoum.Fields("NoLigne") = iCompteurSoum

415           rstImpProjSoum.Fields("IDSoumission") = sNoProjet

420           rstImpProjSoum.Fields("SousSection") = sSousSectionRS
                    
425           Call rstImpProjSoum.Update
          
430           iCompteurSoum = iCompteurSoum + 1
435         End If
440       End If
         
          'ajoute une piece dans impression_soum
445       Call rstImpProjSoum.AddNew
      
450       rstImpProjSoum.Fields("NoLigne") = iCompteurSoum

455       rstImpProjSoum.Fields("IDsoumission") = sNoProjet

460       rstImpProjSoum.Fields("NumItem") = rstPiece.Fields("NumItem")
465       rstImpProjSoum.Fields("Qté") = rstPiece.Fields("Qté")
      
470       If m_eLangage = ANGLAIS Then
475         rstImpProjSoum.Fields("Description") = rstPiece.Fields("DESC_EN")
480       Else
485         rstImpProjSoum.Fields("Description") = rstPiece.Fields("DESC_FR")
490       End If

        '************************************************************************************************
        'SECTION MODIFIÉ PAR GAÉTAN GINGRAS LE 07 FÉVRIER 2010
        '************************************************************************************************
     
495       rstImpProjSoum.Fields("MANUFACT") = rstPiece.Fields("MANUFACT")
500       'rstImpProjSoum.Fields("PRIX_LIST") = rstPiece.Fields("PRIX_LIST")
      
505       'If Trim(rstPiece.Fields("ESCOMPTE")) <> vbNullString Then
510       '  rstImpProjSoum.Fields("ESCOMPTE") = Conversion(Replace(rstPiece.Fields("ESCOMPTE"), ".", ","), MODE_POURCENT)
515       'Else
520       '  rstImpProjSoum.Fields("ESCOMPTE") = Conversion(0, MODE_POURCENT)
525       'End If
     
530       rstImpProjSoum.Fields("PRIX_NET") = rstPiece.Fields("PRIX_NET")
                   
535       Call rstTemp.Open("SELECT NomFournisseur FROM GRB_Fournisseur WHERE IDFRS = " & rstPiece.Fields("IDFRS"), g_connData, adOpenDynamic, adLockOptimistic)

540       If Not rstTemp.EOF Then
545         rstImpProjSoum.Fields("NomFournisseur") = rstTemp.Fields("NomFournisseur")
550       End If
          
555       Call rstTemp.Close
               
560       rstImpProjSoum.Fields("PRIX_TOTAL") = rstPiece.Fields("PRIX_TOTAL")
565       rstImpProjSoum.Fields("PROFIT_ARGENT") = rstPiece.Fields("PROFIT_ARGENT")

        'AJOUT DE CETTE SECTION PAR GAÉTAN GINGRAS LE 07 FÉVRIER 2010
        '*************************************************************
        If m_eType = TYPE_PROJET Then
            If Trim(rstPiece.Fields("DateRéception")) <> vbNullString Then
                rstImpProjSoum.Fields("DateReception") = rstPiece.Fields("DateRéception")
            Else
                rstImpProjSoum.Fields("DateReception") = ""
            End If
            If Trim(rstPiece.Fields("DateCommande")) <> vbNullString Then
                rstImpProjSoum.Fields("DateCommande") = rstPiece.Fields("DateCommande")
            Else
                rstImpProjSoum.Fields("DateCommande") = ""
            End If
        Else    'il n'y a pas de champ date de réception et de commande dans la table GRB_Soumission_Pièces
            rstImpProjSoum.Fields("DateReception") = ""
            rstImpProjSoum.Fields("DateCommande") = ""
        End If
        '************************************************************************************************
        'FIN DE LA SECTION DE MODIFICATION
        '************************************************************************************************

       
570       Call rstImpProjSoum.Update
     
575       iCompteurSoum = iCompteurSoum + 1
    
          'Prochaine enreg
580       Call rstPiece.MoveNext
585     Loop
        
        'Ferme les tables
590     Call rstImpProjSoum.Close
  
        ''''''''''''''''''''''''''''''''''
        ' rapport soumission, met dans l'ordre de ligne
        ''''''''''''''''''''''''''''''''''''
595     Call rstImpProjSoum.Open("SELECT * FROM GRB_impression_soumission WHERE IDSoumission = '" & sNoProjet & "' ORDER BY NoLigne", g_connData, adOpenDynamic, adLockOptimistic)
            
600     Set DR_SoumissionMec.DataSource = rstImpProjSoum

605     Call CalculerPrixFacturation(sNoFacture, sCommission, sPrixTotal, sProfit, sTotalPiece, sImprevue, sTotalTemps, sAutre)
     
610     Call TraduireImpressionSoumission

615     DR_SoumissionMec.Sections("Section5").Controls("shpCadrePrixReception").Visible = False
620     DR_SoumissionMec.Sections("Section5").Controls("lblTitrePrixReception").Visible = False
625     DR_SoumissionMec.Sections("Section5").Controls("lblPrixReception").Visible = False

630     DR_SoumissionMec.Sections("Section5").Controls("shpCadrePrixSoumission").Visible = False
635     DR_SoumissionMec.Sections("Section5").Controls("lblTitrePrixSoumission").Visible = False
640     DR_SoumissionMec.Sections("Section5").Controls("lblPrixSoumission").Visible = False

645     DR_SoumissionMec.Sections("Section5").Controls("shpCadreForfait").Visible = False
650     DR_SoumissionMec.Sections("Section5").Controls("lblTitreForfait").Visible = False
655     DR_SoumissionMec.Sections("Section5").Controls("lblForfait").Visible = False

660     DR_SoumissionMec.Sections("Section2").Controls("lblTitreNoFacture").Visible = True
665     DR_SoumissionMec.Sections("Section2").Controls("lblNoFacture").Visible = True
                             
        'Affiche la date
670     DR_SoumissionMec.Sections("section3").Controls("lbldate").Caption = ConvertDate(Date)
        
        'Affiche entete
675     DR_SoumissionMec.Sections("Section2").Controls("lblSoumission").Caption = rstProjSoum.Fields("IDSoumission")
                  
680     DR_SoumissionMec.Sections("Section2").Controls("lblprojet").Caption = rstProjSoum.Fields("IDProjet")
                
685     DR_SoumissionMec.Sections("Section2").Controls("lbldescription").Caption = rstProjSoum.Fields("Description")
     
690     Call rstTemp.Open("SELECT NomClient FROM GRB_Client WHERE IDClient = " & rstProjSoum.Fields("IDClient"), g_connData, adOpenDynamic, adLockOptimistic)
      
695     DR_SoumissionMec.Sections("Section2").Controls("lblclient").Caption = rstTemp.Fields("NomClient")

700     DR_SoumissionMec.Sections("Section2").Controls("lblNoFacture").Caption = sNoFacture

705     Call rstTemp.Close
      
710     Call rstTemp.Open("SELECT NomContact FROM GRB_Contact WHERE IDContact = " & rstProjSoum.Fields("IDContact"), g_connData, adOpenDynamic, adLockOptimistic)
      
715     DR_SoumissionMec.Sections("Section2").Controls("lblcontact").Caption = rstTemp.Fields("NomContact")
                   
720     Call rstTemp.Close
      
        'Affiche pied d'état
     
        'Temps
725     If m_eType = TYPE_SOUMISSION Then
730       If Not IsNull(rstProjSoum.Fields("TauxDessin")) Then
735         DR_SoumissionMec.Sections("Section5").Controls("lblTauxDessin").Caption = rstProjSoum.Fields("TauxDessin")
740       Else
745         DR_SoumissionMec.Sections("Section5").Controls("lblTauxDessin").Caption = "0"
750       End If

755       If Not IsNull(rstProjSoum.Fields("TauxCoupe")) Then
760         DR_SoumissionMec.Sections("Section5").Controls("lblTauxCoupe").Caption = rstProjSoum.Fields("TauxCoupe")
765       Else
770         DR_SoumissionMec.Sections("Section5").Controls("lblTauxCoupe").Caption = "0"
775       End If

780       If Not IsNull(rstProjSoum.Fields("TauxMachinage")) Then
785         DR_SoumissionMec.Sections("Section5").Controls("lblTauxMachinage").Caption = rstProjSoum.Fields("TauxMachinage")
790       Else
795         DR_SoumissionMec.Sections("Section5").Controls("lblTauxMachinage").Caption = "0"
800       End If

805       If Not IsNull(rstProjSoum.Fields("TauxSoudure")) Then
810         DR_SoumissionMec.Sections("Section5").Controls("lblTauxSoudure").Caption = rstProjSoum.Fields("TauxSoudure")
815       Else
820         DR_SoumissionMec.Sections("Section5").Controls("lblTauxSoudure").Caption = "0"
825       End If

830       If Not IsNull(rstProjSoum.Fields("TauxAssemblage")) Then
835         DR_SoumissionMec.Sections("Section5").Controls("lblTauxAssemblage").Caption = rstProjSoum.Fields("TauxAssemblage")
840       Else
845         DR_SoumissionMec.Sections("Section5").Controls("lblTauxAssemblage").Caption = "0"
850       End If

855       If Not IsNull(rstProjSoum.Fields("TauxPeinture")) Then
860         DR_SoumissionMec.Sections("Section5").Controls("lblTauxPeinture").Caption = rstProjSoum.Fields("TauxPeinture")
865       Else
870         DR_SoumissionMec.Sections("Section5").Controls("lblTauxPeinture").Caption = "0"
875       End If

880       If Not IsNull(rstProjSoum.Fields("TauxTest")) Then
885         DR_SoumissionMec.Sections("Section5").Controls("lblTauxTest").Caption = rstProjSoum.Fields("TauxTest")
890       Else
895         DR_SoumissionMec.Sections("Section5").Controls("lblTauxTest").Caption = "0"
900       End If

905       If Not IsNull(rstProjSoum.Fields("TauxInstallation")) Then
910         DR_SoumissionMec.Sections("Section5").Controls("lblTauxInstallation").Caption = rstProjSoum.Fields("TauxInstallation")
915       Else
920         DR_SoumissionMec.Sections("Section5").Controls("lblTauxInstallation").Caption = "0"
925       End If

930       If Not IsNull(rstProjSoum.Fields("TauxFormation")) Then
935         DR_SoumissionMec.Sections("Section5").Controls("lblTauxFormation").Caption = rstProjSoum.Fields("TauxFormation")
940       Else
945         DR_SoumissionMec.Sections("Section5").Controls("lblTauxFormation").Caption = "0"
950       End If

955       If Not IsNull(rstProjSoum.Fields("TauxGestion")) Then
960         DR_SoumissionMec.Sections("Section5").Controls("lblTauxGestion").Caption = rstProjSoum.Fields("TauxGestion")
965       Else
970         DR_SoumissionMec.Sections("Section5").Controls("lblTauxGestion").Caption = "0"
975       End If

980       If Not IsNull(rstProjSoum.Fields("TauxShipping")) Then
985         DR_SoumissionMec.Sections("Section5").Controls("lblTauxShipping").Caption = rstProjSoum.Fields("TauxShipping")
990       Else
995         DR_SoumissionMec.Sections("Section5").Controls("lblTauxShipping").Caption = "0"
1000      End If

1005      If Not IsNull(rstProjSoum.Fields("TempsDessin")) Then
1010        DR_SoumissionMec.Sections("Section5").Controls("lblTempsDessinSoum").Caption = rstProjSoum.Fields("TempsDessin")
1015      Else
1020        DR_SoumissionMec.Sections("Section5").Controls("lblTempsDessinSoum").Caption = "0"
1025      End If

1030      If Not IsNull(rstProjSoum.Fields("TempsCoupe")) Then
1035        DR_SoumissionMec.Sections("Section5").Controls("lblTempsCoupeSoum").Caption = rstProjSoum.Fields("TempsCoupe")
1040      Else
1045        DR_SoumissionMec.Sections("Section5").Controls("lblTempsCoupeSoum").Caption = "0"
1050      End If

1055      If Not IsNull(rstProjSoum.Fields("TempsMachinage")) Then
1060        DR_SoumissionMec.Sections("Section5").Controls("lblTempsMachinageSoum").Caption = rstProjSoum.Fields("TempsMachinage")
1065      Else
1070        DR_SoumissionMec.Sections("Section5").Controls("lblTempsMachinageSoum").Caption = "0"
1075      End If

1080      If Not IsNull(rstProjSoum.Fields("TempsSoudure")) Then
1085        DR_SoumissionMec.Sections("Section5").Controls("lblTempsSoudureSoum").Caption = rstProjSoum.Fields("TempsSoudure")
1090      Else
1095        DR_SoumissionMec.Sections("Section5").Controls("lblTempsSoudureSoum").Caption = "0"
1100      End If

1105      If Not IsNull(rstProjSoum.Fields("TempsAssemblage")) Then
1110        DR_SoumissionMec.Sections("Section5").Controls("lblTempsAssemblageSoum").Caption = rstProjSoum.Fields("TempsAssemblage")
1115      Else
1120        DR_SoumissionMec.Sections("Section5").Controls("lblTempsAssemblageSoum").Caption = "0"
1125      End If

1130      If Not IsNull(rstProjSoum.Fields("TempsPeinture")) Then
1135        DR_SoumissionMec.Sections("Section5").Controls("lblTempsPeintureSoum").Caption = rstProjSoum.Fields("TempsPeinture")
1140      Else
1145        DR_SoumissionMec.Sections("Section5").Controls("lblTempsPeintureSoum").Caption = "0"
1150      End If

1155      If Not IsNull(rstProjSoum.Fields("TempsTest")) Then
1160        DR_SoumissionMec.Sections("Section5").Controls("lblTempsTestSoum").Caption = rstProjSoum.Fields("TempsTest")
1165      Else
1170        DR_SoumissionMec.Sections("Section5").Controls("lblTempsTestSoum").Caption = "0"
1175      End If

1180      If Not IsNull(rstProjSoum.Fields("TempsInstallation")) Then
1185        DR_SoumissionMec.Sections("Section5").Controls("lblTempsInstallationSoum").Caption = rstProjSoum.Fields("TempsInstallation")
1190      Else
1195        DR_SoumissionMec.Sections("Section5").Controls("lblTempsInstallationSoum").Caption = "0"
1200      End If

1205      If Not IsNull(rstProjSoum.Fields("TempsFormation")) Then
1210        DR_SoumissionMec.Sections("Section5").Controls("lblTempsFormationSoum").Caption = rstProjSoum.Fields("TempsFormation")
1215      Else
1220        DR_SoumissionMec.Sections("Section5").Controls("lblTempsFormationSoum").Caption = "0"
1225      End If

1230      If Not IsNull(rstProjSoum.Fields("TempsGestion")) Then
1235        DR_SoumissionMec.Sections("Section5").Controls("lblTempsGestionSoum").Caption = rstProjSoum.Fields("TempsGestion")
1240      Else
1245        DR_SoumissionMec.Sections("Section5").Controls("lblTempsGestionSoum").Caption = "0"
1250      End If

1255      If Not IsNull(rstProjSoum.Fields("TempsShipping")) Then
1260        DR_SoumissionMec.Sections("Section5").Controls("lblTempsShippingSoum").Caption = rstProjSoum.Fields("TempsShipping")
1265      Else
1270        DR_SoumissionMec.Sections("Section5").Controls("lblTempsShippingSoum").Caption = "0"
1275      End If
  
1280      If Not IsNull(rstProjSoum.Fields("TempsDessin")) Then
1285        If IsNumeric(rstProjSoum.Fields("TempsDessin")) Then
1290          dblTempsDessin = CDbl(rstProjSoum.Fields("TempsDessin"))
1295        Else
1300          dblTempsDessin = 0
1305        End If
1310      Else
1315        dblTempsDessin = 0
1320      End If
  
1325      If Not IsNull(rstProjSoum.Fields("TempsCoupe")) Then
1330        If IsNumeric(rstProjSoum.Fields("TempsCoupe")) Then
1335          dblTempsCoupe = CDbl(rstProjSoum.Fields("TempsCoupe"))
1340        Else
1345          dblTempsCoupe = 0
1350        End If
1355      Else
1360        dblTempsCoupe = 0
1365      End If

1370      If Not IsNull(rstProjSoum.Fields("TempsMachinage")) Then
1375        If IsNumeric(rstProjSoum.Fields("TempsMachinage")) Then
1380          dblTempsMachinage = CDbl(rstProjSoum.Fields("TempsMachinage"))
1385        Else
1390          dblTempsMachinage = 0
1395        End If
1400      Else
1405        dblTempsMachinage = 0
1410      End If

1415      If Not IsNull(rstProjSoum.Fields("TempsSoudure")) Then
1420        If IsNumeric(rstProjSoum.Fields("TempsSoudure")) Then
1425          dblTempsSoudure = CDbl(rstProjSoum.Fields("TempsSoudure"))
1430        Else
1435          dblTempsSoudure = 0
1440        End If
1445      Else
1450        dblTempsSoudure = 0
1455      End If

1460      If Not IsNull(rstProjSoum.Fields("TempsAssemblage")) Then
1465        If IsNumeric(rstProjSoum.Fields("TempsAssemblage")) Then
1470          dblTempsAssemblage = CDbl(rstProjSoum.Fields("TempsAssemblage"))
1475        Else
1480          dblTempsAssemblage = 0
1485        End If
1490      Else
1495        dblTempsAssemblage = 0
1500      End If

1505      If Not IsNull(rstProjSoum.Fields("TempsPeinture")) Then
1510        If IsNumeric(rstProjSoum.Fields("TempsPeinture")) Then
1515          dblTempsPeinture = CDbl(rstProjSoum.Fields("TempsPeinture"))
1520        Else
1525          dblTempsPeinture = 0
1530        End If
1535      Else
1540        dblTempsPeinture = 0
1545      End If

1550      If Not IsNull(rstProjSoum.Fields("TempsTest")) Then
1555        If IsNumeric(rstProjSoum.Fields("TempsTest")) Then
1560          dblTempsTest = CDbl(rstProjSoum.Fields("TempsTest"))
1565        Else
1570          dblTempsTest = 0
1575        End If
1580      Else
1585        dblTempsTest = 0
1590      End If

1595      If Not IsNull(rstProjSoum.Fields("TempsInstallation")) Then
1600        If IsNumeric(rstProjSoum.Fields("TempsInstallation")) Then
1605          dblTempsInstallation = CDbl(rstProjSoum.Fields("TempsInstallation"))
1610        Else
1615          dblTempsInstallation = 0
1620        End If
1625      Else
1630        dblTempsInstallation = 0
1635      End If

1640      If Not IsNull(rstProjSoum.Fields("TempsFormation")) Then
1645        If IsNumeric(rstProjSoum.Fields("TempsFormation")) Then
1650          dblTempsFormation = CDbl(rstProjSoum.Fields("TempsFormation"))
1655        Else
1660          dblTempsFormation = 0
1665        End If
1670      Else
1675        dblTempsFormation = 0
1680      End If

1685      If Not IsNull(rstProjSoum.Fields("TempsGestion")) Then
1690        If IsNumeric(rstProjSoum.Fields("TempsGestion")) Then
1695          dblTempsGestion = CDbl(rstProjSoum.Fields("TempsGestion"))
1700        Else
1705          dblTempsGestion = 0
1710        End If
1715      Else
1720        dblTempsGestion = 0
1725      End If

1730      If Not IsNull(rstProjSoum.Fields("TempsShipping")) Then
1735        If IsNumeric(rstProjSoum.Fields("TempsShipping")) Then
1740          dblTempsShipping = CDbl(rstProjSoum.Fields("TempsShipping"))
1745        Else
1750          dblTempsShipping = 0
1755        End If
1760      Else
1765        dblTempsShipping = 0
1770      End If

1775      dblTotalTemps = dblTempsDessin + _
                          dblTempsCoupe + _
                          dblTempsMachinage + _
                          dblTempsSoudure + _
                          dblTempsAssemblage + _
                          dblTempsPeinture + _
                          dblTempsTest + _
                          dblTempsInstallation + _
                          dblTempsFormation + _
                          dblTempsGestion + _
                          dblTempsShipping
                          
1780      DR_SoumissionMec.Sections("Section5").Controls("lblTotalTempsRHSoum").Caption = dblTotalTemps

1785      DR_SoumissionMec.Sections("Section5").Controls("lblTempsDessinProj").Caption = "---"
1790      DR_SoumissionMec.Sections("Section5").Controls("lblTempsCoupeProj").Caption = "---"
1795      DR_SoumissionMec.Sections("Section5").Controls("lblTempsMachinageProj").Caption = "---"
1800      DR_SoumissionMec.Sections("Section5").Controls("lblTempsSoudureProj").Caption = "---"
1805      DR_SoumissionMec.Sections("Section5").Controls("lblTempsAssemblageProj").Caption = "---"
1810      DR_SoumissionMec.Sections("Section5").Controls("lblTempsPeintureProj").Caption = "---"
1815      DR_SoumissionMec.Sections("Section5").Controls("lblTempsTestProj").Caption = "---"
1820      DR_SoumissionMec.Sections("Section5").Controls("lblTempsInstallationProj").Caption = "---"
1825      DR_SoumissionMec.Sections("Section5").Controls("lblTempsFormationProj").Caption = "---"
1830      DR_SoumissionMec.Sections("Section5").Controls("lblTempsGestionProj").Caption = "---"
1835      DR_SoumissionMec.Sections("Section5").Controls("lblTempsShippingProj").Caption = "---"

1840      DR_SoumissionMec.Sections("Section5").Controls("lblTotalTempsRHProj").Caption = "---"

1845      DR_SoumissionMec.Sections("Section5").Controls("lblTempsDessinConc").Caption = "---"
1850      DR_SoumissionMec.Sections("Section5").Controls("lblTempsCoupeConc").Caption = "---"
1855      DR_SoumissionMec.Sections("Section5").Controls("lblTempsMachinageConc").Caption = "---"
1860      DR_SoumissionMec.Sections("Section5").Controls("lblTempsSoudureConc").Caption = "---"
1865      DR_SoumissionMec.Sections("Section5").Controls("lblTempsAssemblageConc").Caption = "---"
1870      DR_SoumissionMec.Sections("Section5").Controls("lblTempsPeintureConc").Caption = "---"
1875      DR_SoumissionMec.Sections("Section5").Controls("lblTempsTestConc").Caption = "---"
1880      DR_SoumissionMec.Sections("Section5").Controls("lblTempsInstallationConc").Caption = "---"
1885      DR_SoumissionMec.Sections("Section5").Controls("lblTempsFormationConc").Caption = "---"
1890      DR_SoumissionMec.Sections("Section5").Controls("lblTempsGestionConc").Caption = "---"
1895      DR_SoumissionMec.Sections("Section5").Controls("lblTempsShippingConc").Caption = "---"

1900      DR_SoumissionMec.Sections("Section5").Controls("lblTotalTempsRHConc").Caption = "---"
1905    Else
1910      If Not IsNull(rstProjSoum.Fields("TauxDessin")) Then
1915        DR_SoumissionMec.Sections("Section5").Controls("lblTauxDessin").Caption = rstProjSoum.Fields("TauxDessin")
1920      Else
1925        DR_SoumissionMec.Sections("Section5").Controls("lblTauxDessin").Caption = "0"
1930      End If

1935      If Not IsNull(rstProjSoum.Fields("TauxCoupe")) Then
1940        DR_SoumissionMec.Sections("Section5").Controls("lblTauxCoupe").Caption = rstProjSoum.Fields("TauxCoupe")
1945      Else
1950        DR_SoumissionMec.Sections("Section5").Controls("lblTauxCoupe").Caption = "0"
1955      End If

1960      If Not IsNull(rstProjSoum.Fields("TauxMachinage")) Then
1965        DR_SoumissionMec.Sections("Section5").Controls("lblTauxMachinage").Caption = rstProjSoum.Fields("TauxMachinage")
1970      Else
1975        DR_SoumissionMec.Sections("Section5").Controls("lblTauxMachinage").Caption = "0"
1980      End If

1985      If Not IsNull(rstProjSoum.Fields("TauxSoudure")) Then
1990        DR_SoumissionMec.Sections("Section5").Controls("lblTauxSoudure").Caption = rstProjSoum.Fields("TauxSoudure")
1995      Else
2000        DR_SoumissionMec.Sections("Section5").Controls("lblTauxSoudure").Caption = "0"
2005      End If

2010      If Not IsNull(rstProjSoum.Fields("TauxAssemblage")) Then
2015        DR_SoumissionMec.Sections("Section5").Controls("lblTauxAssemblage").Caption = rstProjSoum.Fields("TauxAssemblage")
2020      Else
2025        DR_SoumissionMec.Sections("Section5").Controls("lblTauxAssemblage").Caption = "0"
2030      End If

2035      If Not IsNull(rstProjSoum.Fields("TauxPeinture")) Then
2040        DR_SoumissionMec.Sections("Section5").Controls("lblTauxPeinture").Caption = rstProjSoum.Fields("TauxPeinture")
2045      Else
2050        DR_SoumissionMec.Sections("Section5").Controls("lblTauxPeinture").Caption = "0"
2055      End If

2060      If Not IsNull(rstProjSoum.Fields("TauxTest")) Then
2065        DR_SoumissionMec.Sections("Section5").Controls("lblTauxTest").Caption = rstProjSoum.Fields("TauxTest")
2070      Else
2075        DR_SoumissionMec.Sections("Section5").Controls("lblTauxTest").Caption = "0"
2080      End If

2085      If Not IsNull(rstProjSoum.Fields("TauxInstallation")) Then
2090        DR_SoumissionMec.Sections("Section5").Controls("lblTauxInstallation").Caption = rstProjSoum.Fields("TauxInstallation")
2095      Else
2100        DR_SoumissionMec.Sections("Section5").Controls("lblTauxInstallation").Caption = "0"
2105      End If

2110      If Not IsNull(rstProjSoum.Fields("TauxFormation")) Then
2115        DR_SoumissionMec.Sections("Section5").Controls("lblTauxFormation").Caption = rstProjSoum.Fields("TauxFormation")
2120      Else
2125        DR_SoumissionMec.Sections("Section5").Controls("lblTauxFormation").Caption = "0"
2130      End If

2135      If Not IsNull(rstProjSoum.Fields("TauxGestion")) Then
2140        DR_SoumissionMec.Sections("Section5").Controls("lblTauxGestion").Caption = rstProjSoum.Fields("TauxGestion")
2145      Else
2150        DR_SoumissionMec.Sections("Section5").Controls("lblTauxGestion").Caption = "0"
2155      End If

2160      If Not IsNull(rstProjSoum.Fields("TauxShipping")) Then
2165        DR_SoumissionMec.Sections("Section5").Controls("lblTauxShipping").Caption = rstProjSoum.Fields("TauxShipping")
2170      Else
2175        DR_SoumissionMec.Sections("Section5").Controls("lblTauxShipping").Caption = "0"
2180      End If

2185      If Not IsNull(rstProjSoum.Fields("IDSoumission")) Then
2190        Set rstSoum = New ADODB.Recordset

2195        Call rstSoum.Open("SELECT * FROM GRB_SoumissionMec WHERE IDSoumission = '" & rstProjSoum.Fields("IDSoumission") & "'", g_connData, adOpenDynamic, adLockOptimistic)

2200        If Not rstSoum.EOF Then
2205          If Not IsNull(rstSoum.Fields("TempsDessin")) Then
2210            DR_SoumissionMec.Sections("Section5").Controls("lblTempsDessinSoum").Caption = rstSoum.Fields("TempsDessin")
2215          Else
2220            DR_SoumissionMec.Sections("Section5").Controls("lblTempsDessinSoum").Caption = "0"
2225          End If

2230          If Not IsNull(rstSoum.Fields("TempsCoupe")) Then
2235            DR_SoumissionMec.Sections("Section5").Controls("lblTempsCoupeSoum").Caption = rstSoum.Fields("TempsCoupe")
2240          Else
2245            DR_SoumissionMec.Sections("Section5").Controls("lblTempsCoupeSoum").Caption = "0"
2250          End If

2255          If Not IsNull(rstSoum.Fields("TempsMachinage")) Then
2260            DR_SoumissionMec.Sections("Section5").Controls("lblTempsMachinageSoum").Caption = rstSoum.Fields("TempsMachinage")
2265          Else
2270            DR_SoumissionMec.Sections("Section5").Controls("lblTempsMachinageSoum").Caption = "0"
2275          End If

2280          If Not IsNull(rstSoum.Fields("TempsSoudure")) Then
2285            DR_SoumissionMec.Sections("Section5").Controls("lblTempsSoudureSoum").Caption = rstSoum.Fields("TempsSoudure")
2290          Else
2295            DR_SoumissionMec.Sections("Section5").Controls("lblTempsSoudureSoum").Caption = "0"
2300          End If

2305          If Not IsNull(rstSoum.Fields("TempsAssemblage")) Then
2310            DR_SoumissionMec.Sections("Section5").Controls("lblTempsAssemblageSoum").Caption = rstSoum.Fields("TempsAssemblage")
2315          Else
2320            DR_SoumissionMec.Sections("Section5").Controls("lblTempsAssemblageSoum").Caption = "0"
2325          End If

2330          If Not IsNull(rstSoum.Fields("TempsPeinture")) Then
2335            DR_SoumissionMec.Sections("Section5").Controls("lblTempsPeintureSoum").Caption = rstSoum.Fields("TempsPeinture")
2340          Else
2345            DR_SoumissionMec.Sections("Section5").Controls("lblTempsPeintureSoum").Caption = "0"
2350          End If

2355          If Not IsNull(rstSoum.Fields("TempsTest")) Then
2360            DR_SoumissionMec.Sections("Section5").Controls("lblTempsTestSoum").Caption = rstSoum.Fields("TempsTest")
2365          Else
2370            DR_SoumissionMec.Sections("Section5").Controls("lblTempsTestSoum").Caption = "0"
2375          End If

2380          If Not IsNull(rstSoum.Fields("TempsInstallation")) Then
2385            DR_SoumissionMec.Sections("Section5").Controls("lblTempsInstallationSoum").Caption = rstSoum.Fields("TempsInstallation")
2390          Else
2395            DR_SoumissionMec.Sections("Section5").Controls("lblTempsInstallationSoum").Caption = "0"
2400          End If

2405          If Not IsNull(rstSoum.Fields("TempsFormation")) Then
2410            DR_SoumissionMec.Sections("Section5").Controls("lblTempsFormationSoum").Caption = rstSoum.Fields("TempsFormation")
2415          Else
2420            DR_SoumissionMec.Sections("Section5").Controls("lblTempsFormationSoum").Caption = "0"
2425          End If

2430          If Not IsNull(rstSoum.Fields("TempsGestion")) Then
2435            DR_SoumissionMec.Sections("Section5").Controls("lblTempsGestionSoum").Caption = rstSoum.Fields("TempsGestion")
2440          Else
2445            DR_SoumissionMec.Sections("Section5").Controls("lblTempsGestionSoum").Caption = "0"
2450          End If

2455          If Not IsNull(rstSoum.Fields("TempsShipping")) Then
2460            DR_SoumissionMec.Sections("Section5").Controls("lblTempsShippingSoum").Caption = rstSoum.Fields("TempsShipping")
2465          Else
2470            DR_SoumissionMec.Sections("Section5").Controls("lblTempsShippingSoum").Caption = "0"
2475          End If
  
2480          If Not IsNull(rstSoum.Fields("TempsDessin")) Then
2485            If IsNumeric(rstSoum.Fields("TempsDessin")) Then
2490              dblTempsDessin = CDbl(rstSoum.Fields("TempsDessin"))
2495            Else
2500              dblTempsDessin = 0
2505            End If
2510          Else
2515            dblTempsDessin = 0
2520          End If
  
2525          If Not IsNull(rstSoum.Fields("TempsCoupe")) Then
2530            If IsNumeric(rstSoum.Fields("TempsCoupe")) Then
2535              dblTempsCoupe = CDbl(rstSoum.Fields("TempsCoupe"))
2540            Else
2545              dblTempsCoupe = 0
2550            End If
2555          Else
2560            dblTempsCoupe = 0
2565          End If
  
2570          If Not IsNull(rstSoum.Fields("TempsMachinage")) Then
2575            If IsNumeric(rstSoum.Fields("TempsMachinage")) Then
2580              dblTempsMachinage = CDbl(rstSoum.Fields("TempsMachinage"))
2585            Else
2590              dblTempsMachinage = 0
2595            End If
2600          Else
2605            dblTempsMachinage = 0
2610          End If

2615          If Not IsNull(rstSoum.Fields("TempsSoudure")) Then
2620            If IsNumeric(rstSoum.Fields("TempsSoudure")) Then
2625              dblTempsSoudure = CDbl(rstSoum.Fields("TempsSoudure"))
2630            Else
2635              dblTempsSoudure = 0
2640            End If
2645          Else
2650            dblTempsSoudure = 0
2655          End If

2660          If Not IsNull(rstSoum.Fields("TempsAssemblage")) Then
2665            If IsNumeric(rstSoum.Fields("TempsAssemblage")) Then
2670              dblTempsAssemblage = CDbl(rstSoum.Fields("TempsAssemblage"))
2675            Else
2680              dblTempsAssemblage = 0
2685            End If
2690          Else
2695            dblTempsAssemblage = 0
2700          End If

2705          If Not IsNull(rstSoum.Fields("TempsPeinture")) Then
2710            If IsNumeric(rstSoum.Fields("TempsPeinture")) Then
2715              dblTempsPeinture = CDbl(rstSoum.Fields("TempsPeinture"))
2720            Else
2725              dblTempsPeinture = 0
2730            End If
2735          Else
2740            dblTempsPeinture = 0
2745          End If

2750          If Not IsNull(rstSoum.Fields("TempsTest")) Then
2755            If IsNumeric(rstSoum.Fields("TempsTest")) Then
2760              dblTempsTest = CDbl(rstSoum.Fields("TempsTest"))
2765            Else
2770              dblTempsTest = 0
2775            End If
2780          Else
2785            dblTempsTest = 0
2790          End If

2795          If Not IsNull(rstSoum.Fields("TempsInstallation")) Then
2800            If IsNumeric(rstSoum.Fields("TempsInstallation")) Then
2805              dblTempsInstallation = CDbl(rstSoum.Fields("TempsInstallation"))
2810            Else
2815              dblTempsInstallation = 0
2820            End If
2825          Else
2830            dblTempsInstallation = 0
2835          End If

2840          If Not IsNull(rstSoum.Fields("TempsFormation")) Then
2845            If IsNumeric(rstSoum.Fields("TempsFormation")) Then
2850              dblTempsFormation = CDbl(rstSoum.Fields("TempsFormation"))
2855            Else
2860              dblTempsFormation = 0
2865            End If
2870          Else
2875            dblTempsFormation = 0
2880          End If

2885          If Not IsNull(rstSoum.Fields("TempsGestion")) Then
2890            If IsNumeric(rstSoum.Fields("TempsGestion")) Then
2895              dblTempsGestion = CDbl(rstSoum.Fields("TempsGestion"))
2900            Else
2905              dblTempsGestion = 0
2910            End If
2915          Else
2920            dblTempsGestion = 0
2925          End If
  
2930          If Not IsNull(rstSoum.Fields("TempsShipping")) Then
2935            If IsNumeric(rstSoum.Fields("TempsShipping")) Then
2940              dblTempsShipping = CDbl(rstSoum.Fields("TempsShipping"))
2945            Else
2950              dblTempsShipping = 0
2955            End If
2960          Else
2965            dblTempsShipping = 0
2970          End If
  
  
2975          dblTotalTemps = dblTempsDessin + _
                              dblTempsCoupe + _
                              dblTempsMachinage + _
                              dblTempsSoudure + _
                              dblTempsAssemblage + _
                              dblTempsPeinture + _
                              dblTempsTest + _
                              dblTempsInstallation + _
                              dblTempsFormation + _
                              dblTempsGestion + _
                              dblTempsShipping
                          
2980          DR_SoumissionMec.Sections("Section5").Controls("lblTotalTempsRHSoum").Caption = dblTotalTemps
2985        Else
2990          DR_SoumissionMec.Sections("Section5").Controls("lblTempsDessinSoum").Caption = "---"
2995          DR_SoumissionMec.Sections("Section5").Controls("lblTempsCoupeSoum").Caption = "---"
3000          DR_SoumissionMec.Sections("Section5").Controls("lblTempsMachinageSoum").Caption = "---"
3005          DR_SoumissionMec.Sections("Section5").Controls("lblTempsSoudureSoum").Caption = "---"
3010          DR_SoumissionMec.Sections("Section5").Controls("lblTempsAssemblageSoum").Caption = "---"
3015          DR_SoumissionMec.Sections("Section5").Controls("lblTempsPeintureSoum").Caption = "---"
3020          DR_SoumissionMec.Sections("Section5").Controls("lblTempsTestSoum").Caption = "---"
3025          DR_SoumissionMec.Sections("Section5").Controls("lblTempsInstallationSoum").Caption = "---"
3030          DR_SoumissionMec.Sections("Section5").Controls("lblTempsFormationSoum").Caption = "---"
3035          DR_SoumissionMec.Sections("Section5").Controls("lblTempsGestionSoum").Caption = "---"
3040          DR_SoumissionMec.Sections("Section5").Controls("lblTempsShippingSoum").Caption = "---"

3045          DR_SoumissionMec.Sections("Section5").Controls("lblTotalTempsRHSoum").Caption = "---"
3050        End If

3055        Call rstSoum.Close
3060        Set rstSoum = Nothing
3065      Else
3070        DR_SoumissionMec.Sections("Section5").Controls("lblTempsDessinSoum").Caption = "---"
3075        DR_SoumissionMec.Sections("Section5").Controls("lblTempsCoupeSoum").Caption = "---"
3080        DR_SoumissionMec.Sections("Section5").Controls("lblTempsMachinageSoum").Caption = "---"
3085        DR_SoumissionMec.Sections("Section5").Controls("lblTempsSoudureSoum").Caption = "---"
3090        DR_SoumissionMec.Sections("Section5").Controls("lblTempsAssemblageSoum").Caption = "---"
3095        DR_SoumissionMec.Sections("Section5").Controls("lblTempsPeintureSoum").Caption = "---"
3100        DR_SoumissionMec.Sections("Section5").Controls("lblTempsTestSoum").Caption = "---"
3105        DR_SoumissionMec.Sections("Section5").Controls("lblTempsInstallationSoum").Caption = "---"
3110        DR_SoumissionMec.Sections("Section5").Controls("lblTempsFormationSoum").Caption = "---"
3115        DR_SoumissionMec.Sections("Section5").Controls("lblTempsGestionSoum").Caption = "---"
3120        DR_SoumissionMec.Sections("Section5").Controls("lblTempsShippingSoum").Caption = "---"

3125        DR_SoumissionMec.Sections("Section5").Controls("lblTotalTempsRHSoum").Caption = "---"
3130      End If

3135      If Not IsNull(rstProjSoum.Fields("TempsDessinProj")) Then
3140        DR_SoumissionMec.Sections("Section5").Controls("lblTempsDessinProj").Caption = rstProjSoum.Fields("TempsDessinProj")
3145      Else
3150        DR_SoumissionMec.Sections("Section5").Controls("lblTempsDessinProj").Caption = "0"
3155      End If

3160      If Not IsNull(rstProjSoum.Fields("TempsCoupeProj")) Then
3165        DR_SoumissionMec.Sections("Section5").Controls("lblTempsCoupeProj").Caption = rstProjSoum.Fields("TempsCoupeProj")
3170      Else
3175        DR_SoumissionMec.Sections("Section5").Controls("lblTempsCoupeProj").Caption = "0"
3180      End If

3185      If Not IsNull(rstProjSoum.Fields("TempsMachinageProj")) Then
3190        DR_SoumissionMec.Sections("Section5").Controls("lblTempsMachinageProj").Caption = rstProjSoum.Fields("TempsMachinageProj")
3195      Else
3200        DR_SoumissionMec.Sections("Section5").Controls("lblTempsMachinageProj").Caption = "0"
3205      End If

3210      If Not IsNull(rstProjSoum.Fields("TempsSoudureProj")) Then
3215        DR_SoumissionMec.Sections("Section5").Controls("lblTempsSoudureProj").Caption = rstProjSoum.Fields("TempsSoudureProj")
3220      Else
3225        DR_SoumissionMec.Sections("Section5").Controls("lblTempsSoudureProj").Caption = "0"
3230      End If

3235      If Not IsNull(rstProjSoum.Fields("TempsAssemblageProj")) Then
3240        DR_SoumissionMec.Sections("Section5").Controls("lblTempsAssemblageProj").Caption = rstProjSoum.Fields("TempsAssemblageProj")
3245      Else
3250        DR_SoumissionMec.Sections("Section5").Controls("lblTempsAssemblageProj").Caption = "0"
3255      End If

3260      If Not IsNull(rstProjSoum.Fields("TempsPeintureProj")) Then
3265        DR_SoumissionMec.Sections("Section5").Controls("lblTempsPeintureProj").Caption = rstProjSoum.Fields("TempsPeintureProj")
3270      Else
3275        DR_SoumissionMec.Sections("Section5").Controls("lblTempsPeintureProj").Caption = "0"
3280      End If

3285      If Not IsNull(rstProjSoum.Fields("TempsTestProj")) Then
3290        DR_SoumissionMec.Sections("Section5").Controls("lblTempsTestProj").Caption = rstProjSoum.Fields("TempsTestProj")
3295      Else
3300        DR_SoumissionMec.Sections("Section5").Controls("lblTempsTestProj").Caption = "0"
3305      End If

3310      If Not IsNull(rstProjSoum.Fields("TempsInstallationProj")) Then
3315        DR_SoumissionMec.Sections("Section5").Controls("lblTempsInstallationProj").Caption = rstProjSoum.Fields("TempsInstallationProj")
3320      Else
3325        DR_SoumissionMec.Sections("Section5").Controls("lblTempsInstallationProj").Caption = "0"
3330      End If

3335      If Not IsNull(rstProjSoum.Fields("TempsFormationProj")) Then
3340        DR_SoumissionMec.Sections("Section5").Controls("lblTempsFormationProj").Caption = rstProjSoum.Fields("TempsFormationProj")
3345      Else
3350        DR_SoumissionMec.Sections("Section5").Controls("lblTempsFormationProj").Caption = "0"
3355      End If

3360      If Not IsNull(rstProjSoum.Fields("TempsGestionProj")) Then
3365        DR_SoumissionMec.Sections("Section5").Controls("lblTempsGestionProj").Caption = rstProjSoum.Fields("TempsGestionProj")
3370      Else
3375        DR_SoumissionMec.Sections("Section5").Controls("lblTempsGestionProj").Caption = "0"
3380      End If

3385      If Not IsNull(rstProjSoum.Fields("TempsShippingProj")) Then
3390        DR_SoumissionMec.Sections("Section5").Controls("lblTempsShippingProj").Caption = rstProjSoum.Fields("TempsShippingProj")
3395      Else
3400        DR_SoumissionMec.Sections("Section5").Controls("lblTempsShippingProj").Caption = "0"
3405      End If

3410      If Not IsNull(rstProjSoum.Fields("TempsDessinProj")) Then
3415        If IsNumeric(rstProjSoum.Fields("TempsDessinProj")) Then
3420          dblTempsDessin = CDbl(rstProjSoum.Fields("TempsDessinProj"))
3425        Else
3430          dblTempsDessin = 0
3435        End If
3440      Else
3445        dblTempsDessin = 0
3450      End If
  
3455      If Not IsNull(rstProjSoum.Fields("TempsCoupeProj")) Then
3460        If IsNumeric(rstProjSoum.Fields("TempsCoupeProj")) Then
3465          dblTempsCoupe = CDbl(rstProjSoum.Fields("TempsCoupeProj"))
3470        Else
3475          dblTempsCoupe = 0
3480        End If
3485      Else
3490        dblTempsCoupe = 0
3495      End If

3500      If Not IsNull(rstProjSoum.Fields("TempsMachinageProj")) Then
3505        If IsNumeric(rstProjSoum.Fields("TempsMachinageProj")) Then
3510          dblTempsMachinage = CDbl(rstProjSoum.Fields("TempsMachinageProj"))
3515        Else
3520          dblTempsMachinage = 0
3525        End If
3530      Else
3535        dblTempsMachinage = 0
3540      End If

3545      If Not IsNull(rstProjSoum.Fields("TempsSoudureProj")) Then
3550        If IsNumeric(rstProjSoum.Fields("TempsSoudureProj")) Then
3555          dblTempsSoudure = CDbl(rstProjSoum.Fields("TempsSoudureProj"))
3560        Else
3565          dblTempsSoudure = 0
3570        End If
3575      Else
3580        dblTempsSoudure = 0
3585      End If

3590      If Not IsNull(rstProjSoum.Fields("TempsAssemblageProj")) Then
3595        If IsNumeric(rstProjSoum.Fields("TempsAssemblageProj")) Then
3600          dblTempsAssemblage = CDbl(rstProjSoum.Fields("TempsAssemblageProj"))
3605        Else
3610          dblTempsAssemblage = 0
3615        End If
3620      Else
3625        dblTempsAssemblage = 0
3630      End If

3635      If Not IsNull(rstProjSoum.Fields("TempsPeintureProj")) Then
3640        If IsNumeric(rstProjSoum.Fields("TempsPeintureProj")) Then
3645          dblTempsPeinture = CDbl(rstProjSoum.Fields("TempsPeintureProj"))
3650        Else
3655          dblTempsPeinture = 0
3660        End If
3665      Else
3670        dblTempsPeinture = 0
3675      End If

3680      If Not IsNull(rstProjSoum.Fields("TempsTestProj")) Then
3685        If IsNumeric(rstProjSoum.Fields("TempsTestProj")) Then
3690          dblTempsTest = CDbl(rstProjSoum.Fields("TempsTestProj"))
3695        Else
3700          dblTempsTest = 0
3705        End If
3710      Else
3715        dblTempsTest = 0
3720      End If

3725      If Not IsNull(rstProjSoum.Fields("TempsInstallationProj")) Then
3730        If IsNumeric(rstProjSoum.Fields("TempsInstallationProj")) Then
3735          dblTempsInstallation = CDbl(rstProjSoum.Fields("TempsInstallationProj"))
3740        Else
3745          dblTempsInstallation = 0
3750        End If
3755      Else
3760        dblTempsInstallation = 0
3765      End If

3770      If Not IsNull(rstProjSoum.Fields("TempsFormationProj")) Then
3775        If IsNumeric(rstProjSoum.Fields("TempsFormationProj")) Then
3780          dblTempsFormation = CDbl(rstProjSoum.Fields("TempsFormationProj"))
3785        Else
3790          dblTempsFormation = 0
3795        End If
3800      Else
3805        dblTempsFormation = 0
3810      End If

3815      If Not IsNull(rstProjSoum.Fields("TempsGestionProj")) Then
3820        If IsNumeric(rstProjSoum.Fields("TempsGestionProj")) Then
3825          dblTempsGestion = CDbl(rstProjSoum.Fields("TempsGestionProj"))
3830        Else
3835          dblTempsGestion = 0
3840        End If
3845      Else
3850        dblTempsGestion = 0
3855      End If

3860      If Not IsNull(rstProjSoum.Fields("TempsShippingProj")) Then
3865        If IsNumeric(rstProjSoum.Fields("TempsShippingProj")) Then
3870          dblTempsShipping = CDbl(rstProjSoum.Fields("TempsShippingProj"))
3875        Else
3880          dblTempsShipping = 0
3885        End If
3890      Else
3895        dblTempsShipping = 0
3900      End If

3905      dblTotalTemps = dblTempsDessin + _
                          dblTempsCoupe + _
                          dblTempsMachinage + _
                          dblTempsSoudure + _
                          dblTempsAssemblage + _
                          dblTempsPeinture + _
                          dblTempsTest + _
                          dblTempsInstallation + _
                          dblTempsFormation + _
                          dblTempsGestion + _
                          dblTempsShipping
                          
3910      DR_SoumissionMec.Sections("Section5").Controls("lblTotalTempsRHProj").Caption = dblTotalTemps

3915      If rstProjSoum.Fields("TempsProjBarré") = True Then
3920        If Not IsNull(rstProjSoum.Fields("TempsDessinConc")) Then
3925          DR_SoumissionMec.Sections("Section5").Controls("lblTempsDessinConc").Caption = rstProjSoum.Fields("TempsDessinConc")
3930        Else
3935          DR_SoumissionMec.Sections("Section5").Controls("lblTempsDessinConc").Caption = "0"
3940        End If

3945        If Not IsNull(rstProjSoum.Fields("TempsCoupeConc")) Then
3950          DR_SoumissionMec.Sections("Section5").Controls("lblTempsCoupeConc").Caption = rstProjSoum.Fields("TempsCoupeConc")
3955        Else
3960          DR_SoumissionMec.Sections("Section5").Controls("lblTempsCoupeConc").Caption = "0"
3965        End If

3970        If Not IsNull(rstProjSoum.Fields("TempsMachinageConc")) Then
3975          DR_SoumissionMec.Sections("Section5").Controls("lblTempsMachinageConc").Caption = rstProjSoum.Fields("TempsMachinageConc")
3980        Else
3985          DR_SoumissionMec.Sections("Section5").Controls("lblTempsMachinageConc").Caption = "0"
3990        End If

3995        If Not IsNull(rstProjSoum.Fields("TempsSoudureConc")) Then
4000          DR_SoumissionMec.Sections("Section5").Controls("lblTempsSoudureConc").Caption = rstProjSoum.Fields("TempsSoudureConc")
4005        Else
4010          DR_SoumissionMec.Sections("Section5").Controls("lblTempsSoudureConc").Caption = "0"
4015        End If

4020        If Not IsNull(rstProjSoum.Fields("TempsAssemblageConc")) Then
4025          DR_SoumissionMec.Sections("Section5").Controls("lblTempsAssemblageConc").Caption = rstProjSoum.Fields("TempsAssemblageConc")
4030        Else
4035          DR_SoumissionMec.Sections("Section5").Controls("lblTempsAssemblageConc").Caption = "0"
4040        End If

4045        If Not IsNull(rstProjSoum.Fields("TempsPeintureConc")) Then
4050          DR_SoumissionMec.Sections("Section5").Controls("lblTempsPeintureConc").Caption = rstProjSoum.Fields("TempsPeintureConc")
4055        Else
4060          DR_SoumissionMec.Sections("Section5").Controls("lblTempsPeintureConc").Caption = "0"
4065        End If

4070        If Not IsNull(rstProjSoum.Fields("TempsTestConc")) Then
4075          DR_SoumissionMec.Sections("Section5").Controls("lblTempsTestConc").Caption = rstProjSoum.Fields("TempsTestConc")
4080        Else
4085          DR_SoumissionMec.Sections("Section5").Controls("lblTempsTestConc").Caption = "0"
4090        End If

4095        If Not IsNull(rstProjSoum.Fields("TempsInstallationConc")) Then
4100          DR_SoumissionMec.Sections("Section5").Controls("lblTempsInstallationConc").Caption = rstProjSoum.Fields("TempsInstallationConc")
4105        Else
4110          DR_SoumissionMec.Sections("Section5").Controls("lblTempsInstallationConc").Caption = "0"
4115        End If

4120        If Not IsNull(rstProjSoum.Fields("TempsFormationConc")) Then
4125          DR_SoumissionMec.Sections("Section5").Controls("lblTempsFormationConc").Caption = rstProjSoum.Fields("TempsFormationConc")
4130        Else
4135          DR_SoumissionMec.Sections("Section5").Controls("lblTempsFormationConc").Caption = "0"
4140        End If

4145        If Not IsNull(rstProjSoum.Fields("TempsGestionConc")) Then
4150          DR_SoumissionMec.Sections("Section5").Controls("lblTempsGestionConc").Caption = rstProjSoum.Fields("TempsGestionConc")
4155        Else
4160          DR_SoumissionMec.Sections("Section5").Controls("lblTempsGestionConc").Caption = "0"
4165        End If

4170        If Not IsNull(rstProjSoum.Fields("TempsShippingConc")) Then
4175          DR_SoumissionMec.Sections("Section5").Controls("lblTempsShippingConc").Caption = rstProjSoum.Fields("TempsShippingConc")
4180        Else
4185          DR_SoumissionMec.Sections("Section5").Controls("lblTempsShippingConc").Caption = "0"
4190        End If
  
4195        If Not IsNull(rstProjSoum.Fields("TempsDessinConc")) Then
4200          If IsNumeric(rstProjSoum.Fields("TempsDessinConc")) Then
4205            dblTempsDessin = CDbl(rstProjSoum.Fields("TempsDessinConc"))
4210          Else
4215            dblTempsDessin = 0
4220          End If
4225        Else
4230          dblTempsDessin = 0
4235        End If
  
4240        If Not IsNull(rstProjSoum.Fields("TempsCoupeConc")) Then
4245          If IsNumeric(rstProjSoum.Fields("TempsCoupeConc")) Then
4250            dblTempsCoupe = CDbl(rstProjSoum.Fields("TempsCoupeConc"))
4255          Else
4260            dblTempsCoupe = 0
4265          End If
4270        Else
4275          dblTempsCoupe = 0
4280        End If

4285        If Not IsNull(rstProjSoum.Fields("TempsMachinageConc")) Then
4290          If IsNumeric(rstProjSoum.Fields("TempsMachinageConc")) Then
4295            dblTempsMachinage = CDbl(rstProjSoum.Fields("TempsMachinageConc"))
4300          Else
4305            dblTempsMachinage = 0
4310          End If
4315        Else
4320          dblTempsMachinage = 0
4325        End If

4330        If Not IsNull(rstProjSoum.Fields("TempsSoudureConc")) Then
4335          If IsNumeric(rstProjSoum.Fields("TempsSoudureConc")) Then
4340            dblTempsSoudure = CDbl(rstProjSoum.Fields("TempsSoudureConc"))
4345          Else
4350            dblTempsSoudure = 0
4355          End If
4360        Else
4365          dblTempsSoudure = 0
4370        End If

4375        If Not IsNull(rstProjSoum.Fields("TempsAssemblageConc")) Then
4380          If IsNumeric(rstProjSoum.Fields("TempsAssemblageConc")) Then
4385            dblTempsAssemblage = CDbl(rstProjSoum.Fields("TempsAssemblageConc"))
4390          Else
4395            dblTempsAssemblage = 0
4400          End If
4405        Else
4410          dblTempsAssemblage = 0
4415        End If

4420        If Not IsNull(rstProjSoum.Fields("TempsPeintureConc")) Then
4425          If IsNumeric(rstProjSoum.Fields("TempsPeintureConc")) Then
4430            dblTempsPeinture = CDbl(rstProjSoum.Fields("TempsPeintureConc"))
4435          Else
4440            dblTempsPeinture = 0
4445          End If
4450        Else
4455          dblTempsPeinture = 0
4460        End If

4465        If Not IsNull(rstProjSoum.Fields("TempsTestConc")) Then
4470          If IsNumeric(rstProjSoum.Fields("TempsTestConc")) Then
4475            dblTempsTest = CDbl(rstProjSoum.Fields("TempsTestConc"))
4480          Else
4485            dblTempsTest = 0
4490          End If
4495        Else
4500          dblTempsTest = 0
4505        End If
  
4510        If Not IsNull(rstProjSoum.Fields("TempsInstallationConc")) Then
4515          If IsNumeric(rstProjSoum.Fields("TempsInstallationConc")) Then
4520            dblTempsInstallation = CDbl(rstProjSoum.Fields("TempsInstallationConc"))
4525          Else
4530            dblTempsInstallation = 0
4535          End If
4540        Else
4545          dblTempsInstallation = 0
4550        End If
  
4555        If Not IsNull(rstProjSoum.Fields("TempsFormationConc")) Then
4560          If IsNumeric(rstProjSoum.Fields("TempsFormationConc")) Then
4565            dblTempsFormation = CDbl(rstProjSoum.Fields("TempsFormationConc"))
4570          Else
4575            dblTempsFormation = 0
4580          End If
4585        Else
4590          dblTempsFormation = 0
4595        End If
  
4600        If Not IsNull(rstProjSoum.Fields("TempsGestionConc")) Then
4605          If IsNumeric(rstProjSoum.Fields("TempsGestionConc")) Then
4610            dblTempsGestion = CDbl(rstProjSoum.Fields("TempsGestionConc"))
4615          Else
4620            dblTempsGestion = 0
4625          End If
4630        Else
4635          dblTempsGestion = 0
4640        End If
  
4645        If Not IsNull(rstProjSoum.Fields("TempsShippingConc")) Then
4650          If IsNumeric(rstProjSoum.Fields("TempsShippingConc")) Then
4655            dblTempsShipping = CDbl(rstProjSoum.Fields("TempsShippingConc"))
4660          Else
4665            dblTempsShipping = 0
4670          End If
4675        Else
4680          dblTempsShipping = 0
4685        End If
  
  
4690        dblTotalTemps = dblTempsDessin + _
                            dblTempsCoupe + _
                            dblTempsMachinage + _
                            dblTempsSoudure + _
                            dblTempsAssemblage + _
                            dblTempsPeinture + _
                            dblTempsTest + _
                            dblTempsInstallation + _
                            dblTempsFormation + _
                            dblTempsGestion + _
                            dblTempsShipping
                          
4695        DR_SoumissionMec.Sections("Section5").Controls("lblTotalTempsRHConc").Caption = dblTotalTemps
4700      Else
4705        DR_SoumissionMec.Sections("Section5").Controls("lblTempsDessinConc").Caption = "---"
4710        DR_SoumissionMec.Sections("Section5").Controls("lblTempsCoupeConc").Caption = "---"
4715        DR_SoumissionMec.Sections("Section5").Controls("lblTempsMachinageConc").Caption = "---"
4720        DR_SoumissionMec.Sections("Section5").Controls("lblTempsSoudureConc").Caption = "---"
4725        DR_SoumissionMec.Sections("Section5").Controls("lblTempsAssemblageConc").Caption = "---"
4730        DR_SoumissionMec.Sections("Section5").Controls("lblTempsPeintureConc").Caption = "---"
4735        DR_SoumissionMec.Sections("Section5").Controls("lblTempsTestConc").Caption = "---"
4740        DR_SoumissionMec.Sections("Section5").Controls("lblTempsInstallationConc").Caption = "---"
4745        DR_SoumissionMec.Sections("Section5").Controls("lblTempsFormationConc").Caption = "---"
4750        DR_SoumissionMec.Sections("Section5").Controls("lblTempsGestionConc").Caption = "---"
4755        DR_SoumissionMec.Sections("Section5").Controls("lblTempsShippingConc").Caption = "---"

4760        DR_SoumissionMec.Sections("Section5").Controls("lblTotalTempsRHConc").Caption = "---"
4765      End If
4770    End If

4775    Call CalculerTempsReelsImpression(rstProjSoum.Fields("IDProjet"))

        'Autres frais
4780    If Not IsNull(rstProjSoum.Fields("NbrePersonne")) Then
4785      DR_SoumissionMec.Sections("Section5").Controls("lblNbrePersonne").Caption = rstProjSoum.Fields("NbrePersonne")
4790    Else
4795      DR_SoumissionMec.Sections("Section5").Controls("lblNbrePersonne").Caption = "0"
4800    End If

4805    DR_SoumissionMec.Sections("Section5").Controls("lblTempsHebergement").Caption = rstProjSoum.Fields("TempsHebergement")
4810    DR_SoumissionMec.Sections("Section5").Controls("lblTauxHebergement1").Caption = rstProjSoum.Fields("TauxHebergement1")
4815    DR_SoumissionMec.Sections("Section5").Controls("lblTauxHebergement2").Caption = rstProjSoum.Fields("TauxHebergement2")
4820    DR_SoumissionMec.Sections("Section5").Controls("lblTempsRepas").Caption = rstProjSoum.Fields("TempsRepas")
4825    DR_SoumissionMec.Sections("Section5").Controls("lblTauxRepas").Caption = rstProjSoum.Fields("TauxRepas")
4830    DR_SoumissionMec.Sections("Section5").Controls("lblTempsTransport").Caption = rstProjSoum.Fields("TempsTransport")
4835    DR_SoumissionMec.Sections("Section5").Controls("lblTauxTransport").Caption = rstProjSoum.Fields("TauxTransport")
4840    DR_SoumissionMec.Sections("Section5").Controls("lblTempsUniteMobile").Caption = rstProjSoum.Fields("TempsUniteMobile")
4845    DR_SoumissionMec.Sections("Section5").Controls("lblTauxUniteMobile").Caption = rstProjSoum.Fields("TauxUniteMobile")
4850    DR_SoumissionMec.Sections("Section5").Controls("lblPrixEmballage").Caption = rstProjSoum.Fields("PrixEmballage")
4855    DR_SoumissionMec.Sections("Section5").Controls("lblPrixManuel").Caption = rstProjSoum.Fields("Total_Manuel")

4860    DR_SoumissionMec.Sections("Section5").Controls("lblTotalTempsAR").Caption = Conversion(sTotalTemps, MODE_ARGENT)
4865    DR_SoumissionMec.Sections("Section5").Controls("lblTotalPieceAR").Caption = Conversion(sTotalPiece, MODE_ARGENT)
4870    DR_SoumissionMec.Sections("Section5").Controls("lblProfit").Caption = Conversion(rstProjSoum.Fields("profit") * 100, MODE_POURCENT)
4875    DR_SoumissionMec.Sections("Section5").Controls("lblTotalProfit").Caption = Conversion(sProfit, MODE_ARGENT)
4880    DR_SoumissionMec.Sections("Section5").Controls("lblImprevue").Caption = Conversion(rstProjSoum.Fields("imprevue"), MODE_POURCENT)
4885    DR_SoumissionMec.Sections("Section5").Controls("lblImprevueAR").Caption = Conversion(sImprevue, MODE_ARGENT)
4890    DR_SoumissionMec.Sections("Section5").Controls("lblCommission").Caption = Conversion(rstProjSoum.Fields("commission"), MODE_POURCENT)
4895    DR_SoumissionMec.Sections("Section5").Controls("lblCommissionAR").Caption = Conversion(sCommission, MODE_ARGENT)

4900    If Not IsNull(rstProjSoum.Fields("PrixRéception")) Then
4905      DR_SoumissionMec.Sections("Section5").Controls("lblPrixReception").Caption = Conversion(rstProjSoum.Fields("PrixRéception"), MODE_ARGENT)
4910    Else
4915      DR_SoumissionMec.Sections("Section5").Controls("lblPrixReception").Caption = Conversion(0, MODE_ARGENT)
4920    End If

4925    DR_SoumissionMec.Sections("Section5").Controls("lblAutre").Caption = Conversion(sAutre, MODE_ARGENT)

4930    DR_SoumissionMec.Sections("Section5").Controls("lblGrandTotalAR").Caption = Conversion(sPrixTotal, MODE_ARGENT)

        '************************************************************************************************
        'SECTION MODIFIÉ PAR GAÉTAN GINGRAS LE 07 FÉVRIER 2010
        '************************************************************************************************
        If bFlag = True Then
            DR_SoumissionMec.Sections("Section2").Controls("lbl_DateCommande").Visible = True
            DR_SoumissionMec.Sections("Section2").Controls("lbl_DateReception").Visible = True
            DR_SoumissionMec.Sections("Section1").Controls("txt_DateCommande").Visible = True
            DR_SoumissionMec.Sections("Section1").Controls("txt_DateReception").Visible = True
        Else
            DR_SoumissionMec.Sections("Section2").Controls("lbl_DateCommande").Visible = False
            DR_SoumissionMec.Sections("Section2").Controls("lbl_DateReception").Visible = False
            DR_SoumissionMec.Sections("Section1").Controls("txt_DateCommande").Visible = False
            DR_SoumissionMec.Sections("Section1").Controls("txt_DateReception").Visible = False
        End If
        '************************************************************************************************
        'FIN DE LA SECTION MODIFIÉ
        '************************************************************************************************
             
        'Affiche le rapport soumission
4935    DR_SoumissionMec.Orientation = rptOrientLandscape
    
4940    Call DR_SoumissionMec.Show(vbModal)
         
4945    Call rstImpProjSoum.Close
4950    Set rstImpProjSoum = Nothing

4955    Set rstTemp = Nothing
    
4960    Screen.MousePointer = vbDefault

4965    Exit Sub

AfficherErreur:

4970    woups "frmProjSoumMec", "ImprimerProjSoum", Err, Erl
End Sub

Private Sub ImprimerListePiecesFacturation(ByVal rstProjSoum As ADODB.Recordset, ByVal sNoFacture As String)

5       On Error GoTo AfficherErreur

        'Impression de la feuille de soumission
10      Dim rstPiece            As ADODB.Recordset
15      Dim rstTemp             As ADODB.Recordset
20      Dim rstImpListePiece    As ADODB.Recordset
25      Dim iCompteurPiece      As Integer
30      Dim sSousSection        As String
35      Dim sSousSectionRS      As String
40      Dim sChampsSection      As String
45      Dim sOrdreSection       As String
50      Dim sNoProjet           As String
55      Dim sNoSoumission       As String
      
60      Call g_connData.Execute("DELETE * FROM GRB_impression_listepiece")
      
65      iCompteurPiece = 1

70      Screen.MousePointer = vbHourglass
    
75      Set rstImpListePiece = New ADODB.Recordset
    
80      Call rstImpListePiece.Open("SELECT * FROM GRB_impression_listepiece", g_connData, adOpenDynamic, adLockOptimistic)
   
85      sOrdreSection = vbNullString
            
90      sNoProjet = rstProjSoum.Fields("IDProjet")
95      sNoSoumission = rstProjSoum.Fields("IDSoumission")

100     Set rstPiece = New ADODB.Recordset

105     Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoProjet & "' AND Type = 'M' AND Facturation = '" & sNoFacture & "'", g_connData, adOpenDynamic, adLockOptimistic)

110     Set rstTemp = New ADODB.Recordset
    
115     Do While Not rstPiece.EOF
120       If rstPiece.Fields("Visible") = True Then
125         sSousSectionRS = rstPiece.Fields("SousSection")
       
130         If sSousSectionRS = S_PAS_SOUS_SECTION Then
135           sSousSectionRS = " "
140         End If
      
145         If sOrdreSection <> rstPiece.Fields("OrdreSection") Then
              'remplis la table impression_soumission
              'ajoute seulement la section
150           sOrdreSection = rstPiece.Fields("OrdreSection")
        
155           If m_eLangage = ANGLAIS Then
160             sChampsSection = "NomSectionEN"
165           Else
170             sChampsSection = "NomSectionFR"
175           End If

180           Call rstTemp.Open("SELECT " & sChampsSection & " FROM GRB_SoumProjSectionMec WHERE IDSection = " & rstPiece.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
          
185           Call rstImpListePiece.AddNew
          
190           rstImpListePiece("NoLigne") = iCompteurPiece
195           rstImpListePiece("IDSoumission") = sNoSoumission
          
200           If Not IsNull(rstTemp.Fields(sChampsSection)) Then
205             rstImpListePiece.Fields("NomSection") = rstTemp.Fields(sChampsSection)
210           Else
215             rstImpListePiece.Fields("NomSection") = " "
220           End If
          
225           Call rstImpListePiece.Update
                   
230           iCompteurPiece = iCompteurPiece + 1
          
235           Call rstTemp.Close
          
240           sSousSection = rstPiece.Fields("SousSection")
          
245           If sSousSection = S_PAS_SOUS_SECTION Then
250             sSousSection = " "
255           End If
          
260           Call rstImpListePiece.AddNew
          
265           rstImpListePiece.Fields("NoLigne") = iCompteurPiece
270           rstImpListePiece.Fields("IDSoumission") = sNoSoumission
275           rstImpListePiece.Fields("SousSection") = sSousSection
          
280           Call rstImpListePiece.Update
         
285           iCompteurPiece = iCompteurPiece + 1
290         Else
              'ajoute une soussection dans impression_soum
295           If sSousSection <> sSousSectionRS Then
300             sSousSection = sSousSectionRS
        
305             Call rstImpListePiece.AddNew
          
310             rstImpListePiece.Fields("NoLigne") = iCompteurPiece
315             rstImpListePiece.Fields("IDSoumission") = sNoSoumission
320             rstImpListePiece.Fields("SousSection") = sSousSection
            
325             Call rstImpListePiece.Update
              
330             iCompteurPiece = iCompteurPiece + 1
335           End If
340         End If
          
            'ajoute une piece dans la liste de pièce
345         Call rstImpListePiece.AddNew
      
350         rstImpListePiece.Fields("NoLigne") = iCompteurPiece
355         rstImpListePiece.Fields("IDSoumission") = sNoSoumission
360         rstImpListePiece.Fields("numitem") = rstPiece.Fields("numitem")
365         rstImpListePiece.Fields("qté") = rstPiece.Fields("qté")
       
370         If m_eLangage = ANGLAIS Then
375           rstImpListePiece.Fields("description") = rstPiece.Fields("DESC_EN")
380         Else
385           rstImpListePiece.Fields("description") = rstPiece.Fields("DESC_FR")
390         End If
       
395         rstImpListePiece.Fields("manufact") = rstPiece.Fields("manufact")
          
400         Call rstImpListePiece.Update
        
405         iCompteurPiece = iCompteurPiece + 1
410       End If
    
          'prochaine enreg
415       Call rstPiece.MoveNext
420     Loop

425     Call rstPiece.Close
430     Set rstPiece = Nothing
        
435     Call rstImpListePiece.Close
   
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        ' rapport liste piece, met dans l'ordre de ligne '
        ''''''''''''''''''''''''''''''''''''''''''''''''''
      
440     Call rstImpListePiece.Open("SELECT * FROM GRB_impression_Listepiece WHERE IDSoumission = '" & sNoSoumission & "' ORDER BY NoLigne", g_connData, adOpenDynamic, adLockOptimistic)
     
445     Set DR_Liste_piece.DataSource = rstImpListePiece
        
450     Call TraduireImpressionListePiece

455     DR_Liste_piece.Sections("Section4").Controls("lblTitreNoFacture").Visible = True
460     DR_Liste_piece.Sections("Section4").Controls("lblNoFacture").Visible = True
        
        'affiche la date
465     DR_Liste_piece.Sections("section3").Controls("lbldate").Caption = ConvertDate(Date)
      
        'affiche l'entête
470     DR_Liste_piece.Sections("section4").Controls("lblsoumission").Caption = rstProjSoum.Fields("IDSoumission")
      
475     DR_Liste_piece.Sections("section4").Controls("lblprojet").Caption = rstProjSoum.Fields("IDProjet")
    
480     DR_Liste_piece.Sections("section4").Controls("lbldescription").Caption = rstProjSoum.Fields("Description")
  
485     Call rstTemp.Open("SELECT NomClient FROM GRB_Client WHERE IDClient = " & rstProjSoum.Fields("IDClient"), g_connData, adOpenDynamic, adLockOptimistic)
    
490     DR_Liste_piece.Sections("section4").Controls("lblclient").Caption = rstTemp.Fields("NomClient")

495     DR_Liste_piece.Sections("Section4").Controls("lblNoFacture").Caption = sNoFacture
  
500     Call rstTemp.Close
        
505     Call rstTemp.Open("SELECT NomContact FROM GRB_Contact WHERE IDContact = " & rstProjSoum.Fields("IDContact"), g_connData, adOpenDynamic, adLockOptimistic)
      
510     DR_Liste_piece.Sections("section4").Controls("lblcontact").Caption = rstTemp.Fields("NomContact")
      
515     Call rstTemp.Close
    
        'affiche le rapport liste des pieces
520     DR_Liste_piece.Orientation = rptOrientPortrait
  
525     Call DR_Liste_piece.Show(vbModal)
        
530     Call rstImpListePiece.Close
535     Set rstImpListePiece = Nothing
    
540     Set rstTemp = Nothing
    
545     Screen.MousePointer = vbDefault

550     Exit Sub

AfficherErreur:

555     woups "frmProjSoumMec", "ImprimerListePiecesFacturation", Err, Erl
End Sub

Private Sub CalculerPrixFacturation(ByVal sNoFacturation As String, ByRef sCommission As String, ByRef sPrixTotal As String, ByRef sProfit As String, ByRef sTotalPiece As String, ByRef sImprevue As String, ByRef sTotalTemps As String, ByRef sAutre As String)

5       On Error GoTo AfficherErreur

        'Méthode pour calculer le prix
10      Dim dblPrixPieces        As Double
15      Dim dblPrixTotal         As Double
20      Dim dblCommission        As Double
25      Dim dblTotalTemps        As Double
30      Dim dblProfit            As Double
35      Dim dblTotalManuel       As Double
40      Dim dblTotalImprevue     As Double
45      Dim dblGrandTotal        As Double
50      Dim dblTotalDessin       As Double
55      Dim dblTotalCoupe        As Double
60      Dim dblTotalMachinage    As Double
65      Dim dblTotalSoudure      As Double
70      Dim dblTotalAssemblage   As Double
75      Dim dblTotalPeinture     As Double
80      Dim dblTotalTest         As Double
85      Dim dblTotalInstallation As Double
90      Dim dblTotalFormation    As Double
95      Dim dblTotalGestion      As Double
100     Dim dblTotalShipping     As Double
105     Dim dblHebergement       As Double
110     Dim dblRepas             As Double
115     Dim dblTransport         As Double
120     Dim dblUniteMobile       As Double
125     Dim dblPrixEmballage     As Double
130     Dim dblTotalResteTemps   As Double
135     Dim bDemande             As Boolean
140     Dim iNbrePersonne        As Integer
145     Dim iCompteur            As Integer
        
        'Pour chaque élément du listview
150     For iCompteur = 1 To lvwSoumission.ListItems.count
          'Si ce n'est pas une section
155       If lvwSoumission.ListItems(iCompteur).Tag <> vbNullString Then
            'Si ce n'est pas une sous-section
160         If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> vbNullString And lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> "Texte" And lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> "Text" Then
165           If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_FACTURATION) = sNoFacturation Then
                'On additionne le prix total
170             dblPrixPieces = dblPrixPieces + lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_TOTAL) - lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PROFIT)
        
                'On additionne le profit
175             dblProfit = dblProfit + lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PROFIT)
180           End If
185         End If
190       End If
195     Next
  
        'Total des temps
200     dblTotalDessin = CDbl(m_sTempsDessin) * CDbl(m_sTauxDessin)
205     dblTotalCoupe = CDbl(m_sTempsCoupe) * CDbl(m_sTauxCoupe)
210     dblTotalMachinage = CDbl(m_sTempsMachinage) * CDbl(m_sTauxMachinage)
215     dblTotalSoudure = CDbl(m_sTempsSoudure) * CDbl(m_sTauxSoudure)
220     dblTotalAssemblage = CDbl(m_sTempsAssemblage) * CDbl(m_sTauxAssemblage)
225     dblTotalPeinture = CDbl(m_sTempsPeinture) * CDbl(m_sTauxPeinture)
230     dblTotalTest = CDbl(m_sTempsTest) * CDbl(m_sTauxTest)
235     dblTotalInstallation = CDbl(m_sTempsInstallation) * CDbl(m_sTauxInstallation)
240     dblTotalFormation = CDbl(m_sTempsFormation) * CDbl(m_sTauxFormation)
245     dblTotalGestion = CDbl(m_sTempsGestion) * CDbl(m_sTauxGestion)
250     dblTotalShipping = CDbl(m_sTempsShipping) * CDbl(m_sTauxShipping)
         
255     dblTotalTemps = dblTotalDessin + _
                        dblTotalCoupe + _
                        dblTotalMachinage + _
                        dblTotalSoudure + _
                        dblTotalAssemblage + _
                        dblTotalPeinture + _
                        dblTotalTest + _
                        dblTotalInstallation + _
                        dblTotalFormation + _
                        dblTotalGestion + _
                        dblTotalShipping
      
260     iNbrePersonne = Int(m_sNbrePersonne)
         
265     Do While iNbrePersonne > 0
270       If iNbrePersonne >= 2 Then
275         dblHebergement = dblHebergement + CDbl(m_sTempsHebergement) * CDbl(m_sTauxHebergement2)
            
280         iNbrePersonne = iNbrePersonne - 2
285       Else
290         dblHebergement = dblHebergement + CDbl(m_sTempsHebergement) * CDbl(m_sTauxHebergement1)
           
295         iNbrePersonne = iNbrePersonne - 1
300       End If
305     Loop
    
310     dblRepas = CDbl(m_sTempsRepas) * CDbl(m_sTauxRepas) * CDbl(m_sNbrePersonne)
315     dblTransport = CDbl(m_sTempsTransport) * CDbl(m_sTauxTransport)
320     dblUniteMobile = CDbl(m_sTempsUniteMobile) * CDbl(m_sTauxUniteMobile)

        'Correction d'un bug de Type Incompatible
325     If IsNumeric(m_sPrixEmballage) Then
330       dblPrixEmballage = CDbl(m_sPrixEmballage)
335     Else
340       dblPrixEmballage = 0
345     End If
    
350     dblTotalResteTemps = dblHebergement + dblRepas + dblTransport + dblUniteMobile + dblPrixEmballage
                                                            
355     If IsNumeric(txtPrixManuel.Text) Then
360       dblTotalManuel = CDbl(txtPrixManuel.Text)
365     Else
370       dblTotalManuel = 0
375     End If
                      
380     dblTotalImprevue = (dblPrixPieces + dblProfit) * CDbl(m_sImprevue)
  
385     dblPrixTotal = dblPrixPieces + dblProfit + dblTotalTemps + dblTotalImprevue + dblTotalManuel + dblTotalResteTemps
                      
        'Le calcul de la commission est sur les manuels (Nbre de page * prix par pages * nbre de copies) + (prix des pièces * pourcentage d'imprévus) + total des temps
390     dblCommission = dblPrixTotal * CDbl(m_sCommission)
        
        'Le prix total est le calcul des manuels (Nbre de page * prix par pages * nbre de copies) + (prix des pièces * pourcentage d'imprévus) + total des temps + total de la commission
395     dblGrandTotal = dblPrixTotal + dblCommission
                
        'Format monétaires avec 2 chiffres après la virgule
400     sTotalPiece = Conversion(CStr(Round(dblPrixPieces, 2)), MODE_ARGENT)
405     sTotalTemps = Conversion(CStr(Round(dblTotalTemps, 2)), MODE_ARGENT)
410     sPrixTotal = Conversion(CStr(Round(dblGrandTotal, 2)), MODE_ARGENT)
415     sImprevue = Conversion(CStr(Round(dblTotalImprevue, 2)), MODE_ARGENT)
420     sCommission = Conversion(CStr(Round(dblCommission, 2)), MODE_ARGENT)
425     sProfit = Conversion(CStr(Round(dblProfit, 2)), MODE_ARGENT)
430     sAutre = Conversion(CStr(Round(dblTotalResteTemps, 2)), MODE_ARGENT)

435     Exit Sub

AfficherErreur:

440     woups "frmProjSoumMec", "CalculerPrix", Err, Erl
End Sub

Private Sub CalculerPrixReception()

5       On Error GoTo AfficherErreur

10      Dim dblPrixReception As Double
15      Dim iCompteur        As Integer
20      Dim itmProjet        As ListItem

25      If m_bDroitPrix = True Then
          'Pour chaque ListItems du ListView
30        For iCompteur = 1 To lvwSoumission.ListItems.count
35          Set itmProjet = lvwSoumission.ListItems(iCompteur)
          
            'Si ce n'est pas une section
40          If itmProjet.Tag <> "" Then
              'Si ce n'est pas une sous-section
45            If itmProjet.SubItems(I_COL_SOUM_PIECE) <> "" Then
                'Si c'est pas du texte
50              If itmProjet.SubItems(I_COL_SOUM_PIECE) <> "Texte" And itmProjet.SubItems(I_COL_SOUM_PIECE) <> "Text" Then
                  'Si c'est une réception
55                If itmProjet.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_GRIS Then
                    'On ajoute le montant
60                  If itmProjet.ListSubItems(I_COL_SOUM_FACTURATION).Tag <> "" Then
65                    dblPrixReception = Round(dblPrixReception + (itmProjet.SubItems(I_COL_SOUM_PRIX_NET) * Replace(itmProjet.ListSubItems(I_COL_SOUM_FACTURATION).Tag, "*", "")), 2)
70                  Else
75                    dblPrixReception = Round(dblPrixReception + 0, 2)
80                  End If
85                Else
                    'Si c'est un retour
90                  If itmProjet.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_ROUGE Then
                      'On soustrait le montant
95                    dblPrixReception = Round(dblPrixReception - (itmProjet.SubItems(I_COL_SOUM_PRIX_NET) * Replace(Replace(itmProjet.Text, "-", ""), "*", "")), 2)
100                 End If
105               End If
110             End If
115           End If
120         End If
125       Next

130       txtPrixReception.Text = Conversion(dblPrixReception, MODE_ARGENT)
135     Else
140       txtPrixReception.Text = ""
145     End If

150     Exit Sub

AfficherErreur:

155     woups "frmProjSoumMec", "CalculerPrixReception", Err, Erl
End Sub

Private Sub AnnulerCommande()

5       On Error GoTo AfficherErreur

10      Dim rstProjet     As ADODB.Recordset
15      Dim itmAvant      As ListItem
20      Dim itmAnnulation As ListItem
25      Dim sExtra        As String

30      If Right$(txtNoProjSoum.Text, 2) >= "00" And Right$(txtNoProjSoum.Text, 2) <= "19" Then
35        sExtra = InputBox("Dans quel extra l'annulation de commande doit être faite ? (2 chiffres seulement)")

40        If Len(sExtra) <> 2 Then
45          Call MsgBox("Format incorrect!", vbOKOnly, "Erreur")

50          Exit Sub
55        End If

60        If Not IsNumeric(sExtra) Then
65          Call MsgBox("L'extra doit être numérique!", vbOKOnly, "Erreur")

70          Exit Sub
75        End If

80        If sExtra < 60 Or sExtra > 98 Then
85          Call MsgBox("L'extra doit être entre 60 et 98!", vbOKOnly, "Erreur")

90          Exit Sub
95        End If

100       Set rstProjet = New ADODB.Recordset

105       Call rstProjet.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sExtra & "'", g_connData, adOpenDynamic, adLockOptimistic)

110       If rstProjet.EOF Then
115         Call MsgBox("Le projet " & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sExtra & " n'existe pas!", vbOKOnly, "Erreur")

120         Call rstProjet.Close
125         Set rstProjet = Nothing

130         Exit Sub
135       Else
140         Call rstProjet.Close
145         Set rstProjet = Nothing
150       End If
155     End If

160     Set itmAvant = lvwSoumission.SelectedItem
165     Set itmAnnulation = lvwSoumission.ListItems.Add(itmAvant.Index + 1)

170     itmAnnulation.Checked = itmAvant.Checked
  
        'Quantité
175     itmAnnulation.Text = "-" & itmAvant.Text

        'On met l'id de la section dans le tag du listItem
180     itmAnnulation.Tag = itmAvant.Tag
                                                                                                         
        'No d'item
185     itmAnnulation.SubItems(I_COL_SOUM_PIECE) = itmAvant.SubItems(I_COL_SOUM_PIECE)
   
        'On met le nom de la sous-section dans le tag du no d'item
190     itmAnnulation.ListSubItems(I_COL_SOUM_PIECE).Tag = itmAvant.ListSubItems(I_COL_SOUM_PIECE).Tag
  
        'On met la description en francais dans la colonne et la description en anglais
        'dans le tag
195     itmAnnulation.SubItems(I_COL_SOUM_DESCR) = itmAvant.SubItems(I_COL_SOUM_DESCR)
200     itmAnnulation.ListSubItems(I_COL_SOUM_DESCR).Tag = itmAvant.ListSubItems(I_COL_SOUM_DESCR).Tag
          
        'On met le nom du fabricant dans la colonne et l'ordre de la section dans le tag
205     itmAnnulation.SubItems(I_COL_SOUM_MANUFACT) = itmAvant.SubItems(I_COL_SOUM_MANUFACT)
210     itmAnnulation.ListSubItems(I_COL_SOUM_MANUFACT).Tag = itmAvant.ListSubItems(I_COL_SOUM_MANUFACT).Tag

        'Prix listé
215     itmAnnulation.SubItems(I_COL_SOUM_PRIX_LIST) = itmAvant.SubItems(I_COL_SOUM_PRIX_LIST)

220     itmAnnulation.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = itmAvant.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag
       
225     itmAnnulation.SubItems(I_COL_SOUM_ESCOMPTE) = itmAvant.SubItems(I_COL_SOUM_ESCOMPTE)

230     itmAnnulation.SubItems(I_COL_SOUM_PRIX_NET) = itmAvant.SubItems(I_COL_SOUM_PRIX_NET)
            
        'On met le fournisseur dans la colonne et l'id dans le tag
235     itmAnnulation.SubItems(I_COL_SOUM_DISTRIB) = itmAvant.SubItems(I_COL_SOUM_DISTRIB)
240     itmAnnulation.ListSubItems(I_COL_SOUM_DISTRIB).Tag = itmAvant.ListSubItems(I_COL_SOUM_DISTRIB).Tag
    
        'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
245     itmAnnulation.SubItems(I_COL_SOUM_TOTAL) = "-" & itmAvant.SubItems(I_COL_SOUM_TOTAL)
      
        'Pour le profit, c'est le prix total - (prix net * quantité)
250     itmAnnulation.SubItems(I_COL_SOUM_PROFIT) = "-" & itmAvant.SubItems(I_COL_SOUM_PROFIT)

255     If Right$(txtNoProjSoum.Text, 2) >= "00" And Right$(txtNoProjSoum.Text, 2) <= "19" Then
          'Pour savoir lors de l'enregistremenet qu'il faut le lier avec un extra
260       itmAnnulation.ListSubItems(I_COL_SOUM_PROFIT).Tag = "ANNULATION " & sExtra
265     End If

270     itmAnnulation.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_VERT_FORET
275     itmAnnulation.ListSubItems(I_COL_SOUM_DESCR).ForeColor = COLOR_VERT_FORET
280     itmAnnulation.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = COLOR_VERT_FORET
285     itmAnnulation.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = COLOR_VERT_FORET
290     itmAnnulation.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = COLOR_VERT_FORET
295     itmAnnulation.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = COLOR_VERT_FORET
300     itmAnnulation.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = COLOR_VERT_FORET
305     itmAnnulation.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = COLOR_VERT_FORET
310     itmAnnulation.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = COLOR_VERT_FORET

315     itmAnnulation.ListSubItems(I_COL_SOUM_PIECE).Bold = True
320     itmAnnulation.ListSubItems(I_COL_SOUM_DESCR).Bold = True
325     itmAnnulation.ListSubItems(I_COL_SOUM_DISTRIB).Bold = True
330     itmAnnulation.ListSubItems(I_COL_SOUM_ESCOMPTE).Bold = True
335     itmAnnulation.ListSubItems(I_COL_SOUM_MANUFACT).Bold = True
340     itmAnnulation.ListSubItems(I_COL_SOUM_PRIX_LIST).Bold = True
345     itmAnnulation.ListSubItems(I_COL_SOUM_PRIX_NET).Bold = True
350     itmAnnulation.ListSubItems(I_COL_SOUM_PROFIT).Bold = True
355     itmAnnulation.ListSubItems(I_COL_SOUM_TOTAL).Bold = True

360     itmAvant.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_VERT_FORET
365     itmAvant.ListSubItems(I_COL_SOUM_DESCR).ForeColor = COLOR_VERT_FORET
370     itmAvant.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = COLOR_VERT_FORET
375     itmAvant.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = COLOR_VERT_FORET
380     itmAvant.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = COLOR_VERT_FORET
385     itmAvant.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = COLOR_VERT_FORET
390     itmAvant.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = COLOR_VERT_FORET
395     itmAvant.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = COLOR_VERT_FORET
400     itmAvant.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = COLOR_VERT_FORET

405     If itmAvant.SubItems(I_COL_SOUM_DATE_COMMANDE) <> "" Then
410       itmAvant.ListSubItems(I_COL_SOUM_DATE_COMMANDE).ForeColor = COLOR_VERT_FORET
415     End If

420     itmAvant.ListSubItems(I_COL_SOUM_DATE_REQUISE).ForeColor = COLOR_VERT_FORET

425     itmAvant.ListSubItems(I_COL_SOUM_NOM_COMMANDE).ForeColor = COLOR_VERT_FORET
430     itmAvant.ListSubItems(I_COL_SOUM_NO_SEQUENTIEL).ForeColor = COLOR_VERT_FORET

435     itmAvant.ListSubItems(I_COL_SOUM_PIECE).Bold = True
440     itmAvant.ListSubItems(I_COL_SOUM_DESCR).Bold = True
445     itmAvant.ListSubItems(I_COL_SOUM_DISTRIB).Bold = True
450     itmAvant.ListSubItems(I_COL_SOUM_ESCOMPTE).Bold = True
455     itmAvant.ListSubItems(I_COL_SOUM_MANUFACT).Bold = True
460     itmAvant.ListSubItems(I_COL_SOUM_PRIX_LIST).Bold = True
465     itmAvant.ListSubItems(I_COL_SOUM_PRIX_NET).Bold = True
470     itmAvant.ListSubItems(I_COL_SOUM_PROFIT).Bold = True
475     itmAvant.ListSubItems(I_COL_SOUM_TOTAL).Bold = True

480     If itmAvant.SubItems(I_COL_SOUM_DATE_COMMANDE) <> "" Then
485       itmAvant.ListSubItems(I_COL_SOUM_DATE_COMMANDE).Bold = True
490     End If

495     itmAvant.ListSubItems(I_COL_SOUM_DATE_REQUISE).Bold = True

500     itmAvant.ListSubItems(I_COL_SOUM_NOM_COMMANDE).Bold = True
505     itmAvant.ListSubItems(I_COL_SOUM_NOM_COMMANDE).Bold = True

510     Call lvwSoumission.Refresh

515     Call CalculerPrix

520     Exit Sub

AfficherErreur:

525     woups "frmProjSoumMec", "AnnulerCommande", Err, Erl
End Sub

Private Sub AjouterSuppressionCollection(ByVal iIndex As Integer)

5       On Error GoTo AfficherErreur

10      If lvwSoumission.ListItems(iIndex).SubItems(I_COL_SOUM_PIECE) <> "Texte" And lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) <> "Text" Then
15        Call m_collQteSupp.Add(Replace(lvwSoumission.ListItems(iIndex).Text, "*", ""))
20        Call m_collNoItemSupp.Add(lvwSoumission.ListItems(iIndex).SubItems(I_COL_SOUM_PIECE))
25        Call m_collDateSupp.Add(ConvertDate(Date))
30        Call m_collHeureSupp.Add(Time)
35      End If

40      Exit Sub

AfficherErreur:

45      woups "frmProjSoumMec", "AjouterSuppressionCollection", Err, Erl
End Sub

Private Sub EnregistrerSuppression()

5       On Error GoTo AfficherErreur

10      Dim rstBavard  As ADODB.Recordset
15      Dim rstEmploye As ADODB.Recordset
20      Dim iNoEmploye As Integer
25      Dim iCompteur  As Integer

30      Set rstBavard = New ADODB.Recordset
35      Set rstEmploye = New ADODB.Recordset

40      Call rstBavard.Open("SELECT * FROM GRB_BavardSuppression", g_connData, adOpenDynamic, adLockOptimistic)

45      Call rstEmploye.Open("SELECT noEmploye FROM GRB_Employés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)

50      iNoEmploye = rstEmploye.Fields("noEmploye")

55      Call rstEmploye.Close
60      Set rstEmploye = Nothing

65      If Not m_collNoItemSupp Is Nothing Then
70        For iCompteur = 1 To m_collNoItemSupp.count
75          Call rstBavard.AddNew

80          rstBavard.Fields("IDUser") = iNoEmploye
85          rstBavard.Fields("NoProjsoum") = txtNoProjSoum.Text
90          rstBavard.Fields("Type") = "M"
95          rstBavard.Fields("Qté") = m_collQteSupp(iCompteur)
100         rstBavard.Fields("No Item") = m_collNoItemSupp(iCompteur)
105         rstBavard.Fields("Date") = m_collDateSupp(iCompteur)
110         rstBavard.Fields("Heure") = m_collHeureSupp(iCompteur)

115         Call rstBavard.Update
120       Next
125     End If

130     Call rstBavard.Close
135     Set rstBavard = Nothing

140     Exit Sub

AfficherErreur:

145     woups "frmProjSoumMec", "EnregistrerSuppression", Err, Erl
End Sub

Private Sub RemplirListViewSuppression()

5       On Error GoTo AfficherErreur

        'Rempli le listView avec les pièces supprimées
10      Dim rstBavard  As ADODB.Recordset
15      Dim rstEmploye As ADODB.Recordset
20      Dim itmBavard  As ListItem

25      Call lvwBavard.ListItems.Clear

30      Set rstBavard = New ADODB.Recordset

35      Call rstBavard.Open("SELECT * FROM GRB_BavardSuppression WHERE NoProjSoum = '" & txtNoProjSoum.Text & "' AND Type = 'E' ORDER BY Date, Heure", g_connData, adOpenDynamic, adLockOptimistic)

40      Set rstEmploye = New ADODB.Recordset

45      Do While Not rstBavard.EOF
50        Set itmBavard = lvwBavard.ListItems.Add

55        Call rstEmploye.Open("SELECT Employe FROM GRB_Employés WHERE NoEmploye = " & rstBavard.Fields("IDUser"), g_connData, adOpenDynamic, adLockOptimistic)

60        itmBavard.Text = rstEmploye.Fields("Employe")

65        Call rstEmploye.Close

70        itmBavard.SubItems(I_COL_SUPP_DATE) = rstBavard.Fields("Date")
75        itmBavard.SubItems(I_COL_SUPP_HEURE) = rstBavard.Fields("Heure")
80        itmBavard.SubItems(I_COL_SUPP_QTE) = rstBavard.Fields("Qté")
85        itmBavard.SubItems(I_COL_SUPP_NO_ITEM) = rstBavard.Fields("No Item")

90        Call rstBavard.MoveNext
95      Loop

100     Set rstEmploye = Nothing

105     Call rstBavard.Close
110     Set rstBavard = Nothing

115     Exit Sub

AfficherErreur:

120     woups "frmProjSoumMec", "RemplirListViewSuppression", Err, Erl
End Sub

Private Sub lvwBavard_LostFocus()

5       On Error GoTo AfficherErreur

        'Lorsque le bavard perd le focus, on l'enlève
10      lvwBavard.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMec", "lvwBavard_LostFocus", Err, Erl
End Sub

Private Sub cmdForfait_Click()

5       On Error GoTo AfficherErreur

10      Dim sMontant As String

15      sMontant = InputBox("Quel est le montant du forfait?")

20      If Trim$(sMontant) <> "" Then
25        sMontant = Replace(sMontant, ".", ",")

30        If IsNumeric(sMontant) Then
35          txtForfait.Text = Conversion(sMontant, MODE_ARGENT)

40          lblForfaitInitiale.Caption = g_sInitiale
45        Else
50          Call MsgBox("Montant non numérique!", vbOKOnly, "Erreur")
55        End If
60     End If

65     Exit Sub

AfficherErreur:

70      woups "frmProjSoumMec", "cmdForfait_Click", Err, Erl
End Sub

Private Sub cmdEffacerForfait_Click()

5       On Error GoTo AfficherErreur

10      txtForfait.Text = ""
15      lblForfaitInitiale.Caption = ""

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumMec", "cmdEffacerForfait_Click", Err, Erl
End Sub

Private Sub CopierPiece()

5       On Error GoTo AfficherErreur

10      Dim itmCopier     As ListItem
15      Dim iNbreSelect   As Integer
20      Dim iCompteur     As Integer
25      Dim iNbreSelected As Integer
30      Dim iIndex        As Integer

35      For iCompteur = 1 To lvwSoumission.ListItems.count
40        If lvwSoumission.ListItems(iCompteur).Selected = True Then
45          If lvwSoumission.ListItems(iCompteur).Tag = "" Or lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) = "" Then
50            Call MsgBox("Impossible de copier, la sélection contient une section ou une sous-section!", vbOKOnly, "Erreur")

55            Exit Sub
60          Else
65            iNbreSelected = iNbreSelected + 1
70          End If
75        End If
80      Next

85      Screen.MousePointer = vbHourglass

90      m_iNbreCopie = iNbreSelected

95      ReDim m_arr_tyCopie(0 To iNbreSelected - 1)

100     For iCompteur = 1 To lvwSoumission.ListItems.count
105       If lvwSoumission.ListItems(iCompteur).Selected = True Then
110         Set itmCopier = lvwSoumission.ListItems(iCompteur)

115         m_arr_tyCopie(iIndex).lColor = itmCopier.ListSubItems(I_COL_SOUM_PIECE).ForeColor

120         m_arr_tyCopie(iIndex).bChecked = itmCopier.Checked

125         m_arr_tyCopie(iIndex).sQuantite = itmCopier.Text

130         m_arr_tyCopie(iIndex).sPiece = itmCopier.SubItems(I_COL_SOUM_PIECE)

135         m_arr_tyCopie(iIndex).sDescr = itmCopier.SubItems(I_COL_SOUM_DESCR)
140         m_arr_tyCopie(iIndex).sDescrTag = itmCopier.ListSubItems(I_COL_SOUM_DESCR).Tag

145         m_arr_tyCopie(iIndex).sManufact = itmCopier.SubItems(I_COL_SOUM_MANUFACT)
  
150         m_arr_tyCopie(iIndex).sPrixList = itmCopier.SubItems(I_COL_SOUM_PRIX_LIST)
155         m_arr_tyCopie(iIndex).sPrixListTag = itmCopier.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag

160         m_arr_tyCopie(iIndex).sEscompte = itmCopier.SubItems(I_COL_SOUM_ESCOMPTE)

165         m_arr_tyCopie(iIndex).sPrixNet = itmCopier.SubItems(I_COL_SOUM_PRIX_NET)

170         m_arr_tyCopie(iIndex).sFRS = itmCopier.SubItems(I_COL_SOUM_DISTRIB)
175         m_arr_tyCopie(iIndex).sFRSTag = itmCopier.ListSubItems(I_COL_SOUM_DISTRIB).Tag

180         m_arr_tyCopie(iIndex).sTotal = itmCopier.SubItems(I_COL_SOUM_TOTAL)
 
185         m_arr_tyCopie(iIndex).sProfit = itmCopier.SubItems(I_COL_SOUM_PROFIT)

190         iIndex = iIndex + 1
195       End If
200     Next

205     Screen.MousePointer = vbDefault

210     Exit Sub

AfficherErreur:

215     woups "frmProjSoumMec", "CopierPiece", Err, Erl
End Sub

Private Sub CollerPiece()

5       On Error GoTo AfficherErreur

10      Dim sIDSection     As String
15      Dim sOrdreSection  As String
20      Dim sSousSection   As String
25      Dim itmColler      As ListItem
30      Dim iCompteur      As Integer
35      Dim iIndexSelected As Integer
40      Dim iIndex         As Integer

        'Pour savoir s'il y a quelque chose à coller
45      If m_iNbreCopie = 0 Then
50        Exit Sub
55      End If
  
60      iIndexSelected = lvwSoumission.SelectedItem.Index
  
65      If iIndexSelected >= 3 Then
70        If lvwSoumission.SelectedItem.Tag = vbNullString Then
75          iIndex = iIndexSelected - 1
80        Else
85          If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) = vbNullString Then
90            If lvwSoumission.ListItems(iIndexSelected - 1).Tag = "" Then
95              Call MsgBox("Impossible de coller la pièce entre une section et une sous-section!", vbOKOnly, "Erreur")

100              Exit Sub
105           Else
110             iIndex = iIndexSelected - 1
115           End If
120         Else
125           iIndex = iIndexSelected
130         End If
135       End If
140     Else
145       Call MsgBox("Emplacement incorrect!", vbOKOnly, "Erreur")

150       Exit Sub
155     End If

160     sIDSection = lvwSoumission.ListItems(iIndex).Tag
165     sOrdreSection = lvwSoumission.ListItems(iIndex).ListSubItems(I_COL_SOUM_MANUFACT).Tag
170     sSousSection = lvwSoumission.ListItems(iIndex).ListSubItems(I_COL_SOUM_PIECE).Tag

175     Screen.MousePointer = vbHourglass

180     For iCompteur = 0 To UBound(m_arr_tyCopie)
185       Set itmColler = lvwSoumission.ListItems.Add(iIndexSelected + iCompteur)

190       itmColler.Checked = m_arr_tyCopie(iCompteur).bChecked

195       itmColler.Text = m_arr_tyCopie(iCompteur).sQuantite
200       itmColler.Tag = sIDSection

205       itmColler.SubItems(I_COL_SOUM_PIECE) = m_arr_tyCopie(iCompteur).sPiece
210       itmColler.ListSubItems(I_COL_SOUM_PIECE).Tag = sSousSection

215       itmColler.SubItems(I_COL_SOUM_DESCR) = m_arr_tyCopie(iCompteur).sDescr
220       itmColler.ListSubItems(I_COL_SOUM_DESCR).Tag = m_arr_tyCopie(iCompteur).sDescrTag

225       itmColler.SubItems(I_COL_SOUM_MANUFACT) = m_arr_tyCopie(iCompteur).sManufact
230       itmColler.ListSubItems(I_COL_SOUM_MANUFACT).Tag = sOrdreSection
  
235       itmColler.SubItems(I_COL_SOUM_PRIX_LIST) = m_arr_tyCopie(iCompteur).sPrixList
240       itmColler.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = m_arr_tyCopie(iCompteur).sPrixListTag

245       itmColler.SubItems(I_COL_SOUM_ESCOMPTE) = m_arr_tyCopie(iCompteur).sEscompte

250       itmColler.SubItems(I_COL_SOUM_PRIX_NET) = m_arr_tyCopie(iCompteur).sPrixNet

255       itmColler.SubItems(I_COL_SOUM_DISTRIB) = m_arr_tyCopie(iCompteur).sFRS
260       itmColler.ListSubItems(I_COL_SOUM_DISTRIB).Tag = m_arr_tyCopie(iCompteur).sFRSTag

265       itmColler.SubItems(I_COL_SOUM_TOTAL) = m_arr_tyCopie(iCompteur).sTotal

270       itmColler.SubItems(I_COL_SOUM_PROFIT) = m_arr_tyCopie(iCompteur).sProfit

275       If m_eType = TYPE_PROJET Then
280         itmColler.ListSubItems(I_COL_SOUM_PIECE).ForeColor = m_arr_tyCopie(iCompteur).lColor
285         itmColler.ListSubItems(I_COL_SOUM_DESCR).ForeColor = m_arr_tyCopie(iCompteur).lColor
290         itmColler.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = m_arr_tyCopie(iCompteur).lColor
295         itmColler.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = m_arr_tyCopie(iCompteur).lColor
300         itmColler.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = m_arr_tyCopie(iCompteur).lColor
305         itmColler.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = m_arr_tyCopie(iCompteur).lColor
310         itmColler.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = m_arr_tyCopie(iCompteur).lColor
315         itmColler.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = m_arr_tyCopie(iCompteur).lColor
320         itmColler.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = m_arr_tyCopie(iCompteur).lColor
325       Else
330         itmColler.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_NOIR
335         itmColler.ListSubItems(I_COL_SOUM_DESCR).ForeColor = COLOR_NOIR
340         itmColler.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = COLOR_NOIR
345         itmColler.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = COLOR_NOIR
350         itmColler.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = COLOR_NOIR
355         itmColler.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = COLOR_NOIR
360         itmColler.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = COLOR_NOIR
365         itmColler.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = COLOR_NOIR
370         itmColler.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = COLOR_NOIR
375       End If

380       Call lvwSoumission.Refresh
385     Next

390     Call CalculerPrix

395     Screen.MousePointer = vbDefault

400     Exit Sub

AfficherErreur:

405     woups "frmProjSoumMec", "CollerPiece", Err, Erl
End Sub

Private Sub Deselect()

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer
  
15      If lvwSoumission.ListItems.count > 0 Then
20        For iCompteur = 1 To lvwSoumission.ListItems.count
25          lvwSoumission.ListItems(iCompteur).Selected = False
30        Next
35      End If

40      Exit Sub

AfficherErreur:

45      woups "frmProjSoumMec", "Deselect", Err, Erl
End Sub

Private Sub txtPrixSpecial_Change()

5       On Error GoTo AfficherErreur

        'Quand le contenu du prix spécial change
  
        'Si la longueur du texte écrit est plus grand que 0
10      If Len(txtPrixSpecial.Text) > 0 Then
          'On vide l'escompte, le prix net et on les désactive
15        mskEscompte.Text = vbNullString
20        txtPrixNet.Text = vbNullString
    
25        mskEscompte.Enabled = False
30        txtPrixNet.Enabled = False
35      Else
          'Sinon, on active escompte et prix net
40        mskEscompte.Enabled = True
45        txtPrixNet.Enabled = True
50      End If

55      Exit Sub

AfficherErreur:

60      woups "frmProjSoumMec", "txtPrixSpecial_Change", Err, Erl
End Sub

Private Sub txtPrixSpecial_LostFocus()

5       On Error GoTo AfficherErreur

10      txtPrixSpecial.Text = Replace(txtPrixSpecial.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumMec", "txtPrixSpecial_LostFocus", Err, Erl
End Sub

Private Function ValiderFormatMecanique(ByVal sNoProjSoum As String) As Boolean
        
5       On Error GoTo AfficherErreur

10      If UCase(Left$(sNoProjSoum, 1)) = "M" Then
15        ValiderFormatMecanique = True
20      Else
25        Call MsgBox("Un numéro mécanique doit absolument commencé par 'M' !", vbOKOnly, "Erreur")

30        ValiderFormatMecanique = False
35      End If

40      Exit Function

AfficherErreur:

45      woups "FrmProjSoumMec", "ValiderFormatMecanique", Err, Erl
End Function

Private Function ValiderFormatSoumission(ByVal sNoSoumission As String) As Boolean
        
5       On Error GoTo AfficherErreur

10      If Mid$(sNoSoumission, 3, 1) = "1" Then
15        ValiderFormatSoumission = True
20      Else
25        Call MsgBox("Une soumission doit absolument avoir '1' comme 3e caractère !", vbOKOnly, "Erreur")

30        ValiderFormatSoumission = False
35      End If

40      Exit Function

AfficherErreur:

45      woups "FrmProjSoumMec", "ValiderFormatSoumission", Err, Erl
End Function

Private Function ValiderFormatJobSansSoum(ByVal sNoProjet As String) As Boolean
        
5       On Error GoTo AfficherErreur
 
10      If Mid$(sNoProjet, 3, 1) <> "3" And Mid$(sNoProjet, 3, 1) <> "1" Then
15        ValiderFormatJobSansSoum = True
20      Else
25        Call MsgBox("Un projet créé sans soumission ne doit pas être '" & Left$(sNoProjet, 3) & "' !", vbOKOnly, "Erreur")

30        ValiderFormatJobSansSoum = False
35      End If

40      Exit Function

AfficherErreur:

45      woups "FrmProjSoumMec", "ValiderFormatJobSansSoum", Err, Erl
End Function

Private Function ValiderFormatJobAvecSoum(ByVal sNoProjet As String) As Boolean
        
5       On Error GoTo AfficherErreur

10      If Mid$(sNoProjet, 3, 1) = "3" Then
15        ValiderFormatJobAvecSoum = True
20      Else
25        Call MsgBox("Un projet créé à partir d'une soumission doit absolument avec un '3' comme 3e caractère!", vbOKOnly, "Errreu")

30        ValiderFormatJobAvecSoum = False
35      End If

40      Exit Function

AfficherErreur:

45      woups "FrmProjSoumMec", "ValiderFormatJobAvecSoum", Err, Erl
End Function

Private Function ValiderFormatJobExtra(ByVal sNoProjet As String) As Boolean
        
5       On Error GoTo AfficherErreur

10      If CInt(Right$(sNoProjet, 2)) >= 50 And CInt(Right$(sNoProjet, 2)) <= 98 Then
15        ValiderFormatJobExtra = True
20      Else
25        Call MsgBox("L'entension d'un extra doit être compris entre 50 et 98 !", vbOKOnly, "Erreur")

30        ValiderFormatJobExtra = False
35      End If

40      Exit Function

AfficherErreur:

45      woups "FrmProjSoumMec", "ValiderFormatJobExtra", Err, Erl
End Function

Private Sub AjouterProjetAuCumulatif()

5       On Error GoTo AfficherErreur

10      Dim sNoCumulatif          As String
15      Dim rstProj               As ADODB.Recordset
20      Dim rstPieces             As ADODB.Recordset
25      Dim rstProjCumulatif      As ADODB.Recordset
30      Dim rstPiecesCumulatif    As ADODB.Recordset
35      Dim rstProjSoum           As ADODB.Recordset
40      Dim rstEmploye            As ADODB.Recordset
45      Dim rstSoum               As ADODB.Recordset
50      Dim bCumulatifExiste      As Boolean
55      Dim dblNbreManuel         As Double
60      Dim dblPrixEmballage      As Double
65      Dim dblTotalManuel        As Double
70      Dim dblForfait            As Double

75      sNoCumulatif = Left$(txtNoProjSoum.Text, 7) & "99"

80      Set rstProj = New ADODB.Recordset
85      Set rstPieces = New ADODB.Recordset
90      Set rstProjCumulatif = New ADODB.Recordset
95      Set rstPiecesCumulatif = New ADODB.Recordset

100     Call rstProjCumulatif.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = '" & sNoCumulatif & "'", g_connData, adOpenDynamic, adLockOptimistic)

105     If rstProjCumulatif.EOF Then
110       bCumulatifExiste = False

115       Call rstProjCumulatif.AddNew

120       rstProjCumulatif.Fields("IDProjet") = sNoCumulatif

          'Ouverture du projet -01 pour voir la soumission reliée pour ensuite assigner
          'la soumission -99 avec le projet -99
125       Call rstProj.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, 6) & "-01'", g_connData, adOpenForwardOnly, adLockReadOnly)

130       If Not rstProj.EOF Then
135         If Not IsNull(rstProj.Fields("IDSoumission")) Then
140           If Len(rstProj.Fields("IDSoumission")) >= 6 Then
145             Set rstSoum = New ADODB.Recordset

150             Call rstSoum.Open("SELECT * FROM GRB_SoumissionMec WHERE IDSoumission = '" & Left$(rstProj.Fields("IDSoumission"), 6) & "-99'", g_connData, adOpenForwardOnly, adLockReadOnly)

155             If Not rstSoum.EOF Then
160               rstProjCumulatif.Fields("IDSoumission") = rstSoum.Fields("IDSoumission")
165             End If

170             Call rstSoum.Close
175             Set rstSoum = Nothing
180           End If
185         End If
190       End If

195       Call rstProj.Close

200       Call rstProj.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

205       rstProjCumulatif.Fields("IDClient") = rstProj.Fields("IDClient")
210       rstProjCumulatif.Fields("IDContact") = rstProj.Fields("IDContact")

215       rstProjCumulatif.Fields("TauxMachinage") = rstProj.Fields("TauxMachinage")
220       rstProjCumulatif.Fields("TauxCoupe") = rstProj.Fields("TauxCoupe")
225       rstProjCumulatif.Fields("TauxSoudure") = rstProj.Fields("TauxSoudure")
230       rstProjCumulatif.Fields("TauxAssemblage") = rstProj.Fields("TauxAssemblage")
235       rstProjCumulatif.Fields("TauxPeinture") = rstProj.Fields("TauxPeinture")
240       rstProjCumulatif.Fields("TauxTest") = rstProj.Fields("TauxTest")
245       rstProjCumulatif.Fields("TauxDessin") = rstProj.Fields("TauxDessin")
250       rstProjCumulatif.Fields("TauxFormation") = rstProj.Fields("TauxFormation")
255       rstProjCumulatif.Fields("TauxInstallation") = rstProj.Fields("TauxInstallation")
260       rstProjCumulatif.Fields("TauxGestion") = rstProj.Fields("TauxGestion")
265       rstProjCumulatif.Fields("TauxShipping") = rstProj.Fields("TauxShipping")

270       rstProjCumulatif.Fields("Profit") = rstProj.Fields("Profit")
275       rstProjCumulatif.Fields("imprevue") = rstProj.Fields("imprevue")
280       rstProjCumulatif.Fields("commission") = rstProj.Fields("commission")

285       Call rstProj.Close

290       rstProjCumulatif.Fields("Description") = "Cumulatif de " & Left$(txtNoProjSoum.Text, 6)

295       Set rstEmploye = New ADODB.Recordset

300       Call rstEmploye.Open("SELECT * FROM GRB_employés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

305       rstProjCumulatif.Fields("creer") = ConvertDate(Date)

310       rstProjCumulatif.Fields("creer_par") = rstEmploye.Fields("NoEmploye")

315       Call rstEmploye.Close
320       Set rstEmploye = Nothing

325       Call rstProjCumulatif.Update

330       Set rstProjSoum = New ADODB.Recordset

335       Call rstProjSoum.Open("SELECT * FROM GRB_ProjSoum", g_connData, adOpenDynamic, adLockOptimistic)

340       Call rstProjSoum.AddNew

345       rstProjSoum.Fields("IDProjSoum") = sNoCumulatif
350       rstProjSoum.Fields("NoClient") = rstProjCumulatif.Fields("IDClient")
355       rstProjSoum.Fields("Description") = rstProjCumulatif.Fields("Description")
360       rstProjSoum.Fields("DateOuverture") = ConvertDate(Date)
365       rstProjSoum.Fields("Ouvert") = True
370       rstProjSoum.Fields("Verrouillé") = True
375       rstProjSoum.Fields("Type") = "P"

380       Call rstProjSoum.Update
    
385       Call rstProjSoum.Close
390       Set rstProjSoum = Nothing
395     Else
400       bCumulatifExiste = True
405     End If

410     rstProj.CursorLocation = adUseClient

415     Call rstProj.Open("SELECT * FROM GRB_ProjetMec WHERE LEFT(IDProjet, 6) = '" & Left$(txtNoProjSoum.Text, 6) & "' AND RIGHT(IDProjet, 2) <> '99'", g_connData, adOpenForwardOnly, adLockReadOnly)

420     If rstProj.RecordCount = 1 Then
425       rstProjCumulatif.Fields("manuel") = rstProj.Fields("manuel")

430       rstProjCumulatif.Fields("PrixEmballage") = rstProj.Fields("PrixEmballage")

435       rstProjCumulatif.Fields("total_manuel") = rstProj.Fields("total_manuel")

440       rstProjCumulatif.Fields("MontantForfait") = rstProj.Fields("MontantForfait")
445     Else
450       Do While Not rstProj.EOF
455         If Not IsNull(rstProj.Fields("manuel")) Then
460           dblNbreManuel = dblNbreManuel + CDbl(rstProj.Fields("manuel"))
465         End If

470         If Not IsNull(rstProj.Fields("PrixEmballage")) Then
475           dblPrixEmballage = dblPrixEmballage + CDbl(rstProj.Fields("PrixEmballage"))
480         End If

485         If Not IsNull(rstProj.Fields("total_manuel")) Then
490           dblTotalManuel = dblTotalManuel + CDbl(rstProj.Fields("total_manuel"))
495         End If

500         If Not IsNull(rstProj.Fields("MontantForfait")) Then
505           If IsNumeric(rstProj.Fields("MontantForfait")) Then
510             dblForfait = dblForfait + CDbl(rstProj.Fields("MontantForfait"))
515           End If
520         End If

525         Call rstProj.MoveNext
530       Loop

535       rstProjCumulatif.Fields("manuel") = dblNbreManuel
540       rstProjCumulatif.Fields("PrixEmballage") = dblPrixEmballage
545       rstProjCumulatif.Fields("total_manuel") = dblTotalManuel
550       rstProjCumulatif.Fields("MontantForfait") = dblForfait
555     End If

560     Call rstProj.Close

565     Call rstProjCumulatif.Update

570     Call rstProjCumulatif.Close

        'AJOUT DES PIÈCES
575     Call rstPiecesCumulatif.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoCumulatif & "'", g_connData, adOpenDynamic, adLockOptimistic)

580     If bCumulatifExiste = True Then
585       Call g_connData.Execute("DELETE * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoCumulatif & "' AND Provenance = '" & Right$(txtNoProjSoum.Text, 2) & "'")

590       Call rstPieces.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & txtNoProjSoum.Text & "' AND Provenance Is Null OR Provenance = '' ORDER BY NuméroLigne", g_connData, adOpenForwardOnly, adLockReadOnly)
595     Else
600       Call rstPieces.Open("SELECT * FROM GRB_Projet_Pieces WHERE LEFT(IDProjet, 6) = '" & Left$(txtNoProjSoum.Text, 6) & "' AND RIGHT(IDProjet, 2) <> '99' AND Provenance Is Null OR Provenance = '' ORDER BY CInt(OrdreSection), NuméroLigne", g_connData, adOpenForwardOnly, adLockReadOnly)
605     End If

610     Do While Not rstPieces.EOF
615       Call rstPiecesCumulatif.AddNew

620       rstPiecesCumulatif.Fields("IDProjet") = sNoCumulatif
625       rstPiecesCumulatif.Fields("IDSection") = rstPieces.Fields("IDSection")
630       rstPiecesCumulatif.Fields("NumItem") = rstPieces.Fields("NumItem")
635       rstPiecesCumulatif.Fields("Qté") = rstPieces.Fields("Qté")
640       rstPiecesCumulatif.Fields("Desc_FR") = rstPieces.Fields("Desc_FR")
645       rstPiecesCumulatif.Fields("Desc_EN") = rstPieces.Fields("Desc_EN")
650       rstPiecesCumulatif.Fields("Manufact") = rstPieces.Fields("Manufact")
655       rstPiecesCumulatif.Fields("Prix_list") = rstPieces.Fields("Prix_list")
660       rstPiecesCumulatif.Fields("Escompte") = rstPieces.Fields("Escompte")
665       rstPiecesCumulatif.Fields("Prix_net") = rstPieces.Fields("Prix_net")
670       rstPiecesCumulatif.Fields("IDFRS") = rstPieces.Fields("IDFRS")
675       rstPiecesCumulatif.Fields("Prix_total") = rstPieces.Fields("Prix_total")
680       rstPiecesCumulatif.Fields("Profit_Argent") = rstPieces.Fields("Profit_Argent")
685       rstPiecesCumulatif.Fields("SousSection") = rstPieces.Fields("SousSection")
690       rstPiecesCumulatif.Fields("OrdreSection") = rstPieces.Fields("OrdreSection")
695       rstPiecesCumulatif.Fields("NuméroLigne") = rstPieces.Fields("NuméroLigne")
700       rstPiecesCumulatif.Fields("PrixOrigine") = rstPieces.Fields("PrixOrigine")
705       rstPiecesCumulatif.Fields("Type") = rstPieces.Fields("Type")
710       rstPiecesCumulatif.Fields("Visible") = rstPieces.Fields("Visible")
715       rstPiecesCumulatif.Fields("Commentaire") = rstPieces.Fields("Commentaire")
720       rstPiecesCumulatif.Fields("Devise") = rstPieces.Fields("Devise")
725       rstPiecesCumulatif.Fields("Provenance") = Right(rstPieces.Fields("IDProjet"), 2)

730       Call rstPiecesCumulatif.Update

735       Call rstPieces.MoveNext
740     Loop

745     Call rstPiecesCumulatif.Close
750     Call rstPieces.Close

755     Set rstProj = Nothing
760     Set rstPieces = Nothing
765     Set rstProjCumulatif = Nothing
770     Set rstPiecesCumulatif = Nothing

775     Call CalculerTotalRecordset(sNoCumulatif)

780     If bCumulatifExiste = False Then
785       If cmbOuvertFerme.ListIndex = I_CMB_OUVERT Then
790         Call RemplirComboProjSoum(txtNoProjSoum.Text)
795       End If
800     End If

805     Exit Sub

AfficherErreur:

810     woups "FrmProjSoumElec", "AjouterProjetAuCumulatif", Err, Erl
End Sub

Private Sub AjouterSoumissionAuCumulatif()

5       On Error GoTo AfficherErreur

10      Dim sNoCumulatif          As String
15      Dim rstSoum               As ADODB.Recordset
20      Dim rstPieces             As ADODB.Recordset
25      Dim rstSoumCumulatif      As ADODB.Recordset
30      Dim rstPiecesCumulatif    As ADODB.Recordset
35      Dim rstProjSoum           As ADODB.Recordset
40      Dim rstEmploye            As ADODB.Recordset
45      Dim bCumulatifExiste      As Boolean
50      Dim dblNbreManuel         As Double
55      Dim dblTempsMachinage     As Double
60      Dim dblTempsCoupe         As Double
65      Dim dblTempsSoudure       As Double
70      Dim dblTempsAssemblage    As Double
75      Dim dblTempsPeinture      As Double
80      Dim dblTempsTest          As Double
85      Dim dblTempsDessin        As Double
90      Dim dblTempsFormation     As Double
95      Dim dblTempsInstallation  As Double
100     Dim dblTempsGestion       As Double
105     Dim dblTempsShipping      As Double
110     Dim dblTempsTransport     As Double
115     Dim dblTempsUniteMobile   As Double
120     Dim dblTotalHebergement   As Double
125     Dim dblTotalRepas         As Double
130     Dim dblPrixEmballage      As Double
135     Dim dblTotalManuel        As Double
140     Dim dblForfait            As Double

145     sNoCumulatif = Left$(txtNoProjSoum.Text, 7) & "99"

150     Set rstSoum = New ADODB.Recordset
155     Set rstPieces = New ADODB.Recordset
160     Set rstSoumCumulatif = New ADODB.Recordset
165     Set rstPiecesCumulatif = New ADODB.Recordset

170     Call rstSoumCumulatif.Open("SELECT * FROM GRB_SoumissionMec WHERE IDSoumission = '" & sNoCumulatif & "'", g_connData, adOpenDynamic, adLockOptimistic)

175     If rstSoumCumulatif.EOF Then
180       bCumulatifExiste = False

185       Call rstSoumCumulatif.AddNew

190       rstSoumCumulatif.Fields("IDSoumission") = sNoCumulatif

195       Call rstSoum.Open("SELECT * FROM GRB_SoumissionMec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

200       rstSoumCumulatif.Fields("IDClient") = rstSoum.Fields("IDClient")
205       rstSoumCumulatif.Fields("IDContact") = rstSoum.Fields("IDContact")

210       rstSoumCumulatif.Fields("TauxMachinage") = rstSoum.Fields("TauxMachinage")
215       rstSoumCumulatif.Fields("TauxCoupe") = rstSoum.Fields("TauxCoupe")
220       rstSoumCumulatif.Fields("TauxSoudure") = rstSoum.Fields("TauxSoudure")
225       rstSoumCumulatif.Fields("TauxAssemblage") = rstSoum.Fields("TauxAssemblage")
230       rstSoumCumulatif.Fields("TauxPeinture") = rstSoum.Fields("TauxPeinture")
235       rstSoumCumulatif.Fields("TauxTest") = rstSoum.Fields("TauxTest")
240       rstSoumCumulatif.Fields("TauxDessin") = rstSoum.Fields("TauxDessin")
245       rstSoumCumulatif.Fields("TauxFormation") = rstSoum.Fields("TauxFormation")
250       rstSoumCumulatif.Fields("TauxInstallation") = rstSoum.Fields("TauxInstallation")
255       rstSoumCumulatif.Fields("TauxGestion") = rstSoum.Fields("TauxGestion")
260       rstSoumCumulatif.Fields("TauxShipping") = rstSoum.Fields("TauxShipping")

265       rstSoumCumulatif.Fields("TauxHebergement1") = rstSoum.Fields("TauxHebergement1")
270       rstSoumCumulatif.Fields("TauxHebergement2") = rstSoum.Fields("TauxHebergement2")
275       rstSoumCumulatif.Fields("TauxRepas") = rstSoum.Fields("TauxRepas")
280       rstSoumCumulatif.Fields("TauxTransport") = rstSoum.Fields("TauxTransport")
285       rstSoumCumulatif.Fields("TauxUniteMobile") = rstSoum.Fields("TauxUniteMobile")

290       rstSoumCumulatif.Fields("Profit") = rstSoum.Fields("Profit")
295       rstSoumCumulatif.Fields("imprevue") = rstSoum.Fields("imprevue")
300       rstSoumCumulatif.Fields("commission") = rstSoum.Fields("commission")

305       Call rstSoum.Close

310       rstSoumCumulatif.Fields("Description") = "Cumulatif de " & Left$(txtNoProjSoum.Text, 6)

315       Set rstEmploye = New ADODB.Recordset

320       Call rstEmploye.Open("SELECT * FROM GRB_employés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

325       rstSoumCumulatif.Fields("creer") = ConvertDate(Date)

330       rstSoumCumulatif.Fields("creer_par") = rstEmploye.Fields("NoEmploye")

335       Call rstEmploye.Close
340       Set rstEmploye = Nothing

345       Call rstSoumCumulatif.Update

350       Set rstProjSoum = New ADODB.Recordset

355       Call rstProjSoum.Open("SELECT * FROM GRB_ProjSoum", g_connData, adOpenDynamic, adLockOptimistic)

360       Call rstProjSoum.AddNew

365       rstProjSoum.Fields("IDProjSoum") = sNoCumulatif
370       rstProjSoum.Fields("NoClient") = rstSoumCumulatif.Fields("IDClient")
375       rstProjSoum.Fields("Description") = rstSoumCumulatif.Fields("Description")
380       rstProjSoum.Fields("DateOuverture") = ConvertDate(Date)
385       rstProjSoum.Fields("Ouvert") = True
390       rstProjSoum.Fields("Verrouillé") = True
395       rstProjSoum.Fields("Type") = "S"

400       Call rstProjSoum.Update
    
405       Call rstProjSoum.Close
410       Set rstProjSoum = Nothing
415     Else
420       bCumulatifExiste = True
425     End If
     
430     rstSoum.CursorLocation = adUseClient
     
435     Call rstSoum.Open("SELECT * FROM GRB_SoumissionMec WHERE LEFT(IDSoumission, 6) = '" & Left$(txtNoProjSoum.Text, 6) & "' AND RIGHT(IDSoumission, 2) <> '99'", g_connData, adOpenForwardOnly, adLockReadOnly)

440     If rstSoum.RecordCount = 1 Then
445       rstSoumCumulatif.Fields("manuel") = rstSoum.Fields("manuel")

450       rstSoumCumulatif.Fields("TempsMachinage") = rstSoum.Fields("TempsMachinage")
455       rstSoumCumulatif.Fields("TempsCoupe") = rstSoum.Fields("TempsCoupe")
460       rstSoumCumulatif.Fields("TempsSoudure") = rstSoum.Fields("TempsSoudure")
465       rstSoumCumulatif.Fields("TempsAssemblage") = rstSoum.Fields("TempsAssemblage")
470       rstSoumCumulatif.Fields("TempsPeinture") = rstSoum.Fields("TempsPeinture")
475       rstSoumCumulatif.Fields("TempsTest") = rstSoum.Fields("TempsTest")
480       rstSoumCumulatif.Fields("TempsDessin") = rstSoum.Fields("TempsDessin")
485       rstSoumCumulatif.Fields("TempsFormation") = rstSoum.Fields("TempsFormation")
490       rstSoumCumulatif.Fields("TempsInstallation") = rstSoum.Fields("TempsInstallation")
495       rstSoumCumulatif.Fields("TempsGestion") = rstSoum.Fields("TempsGestion")
500       rstSoumCumulatif.Fields("TempsShipping") = rstSoum.Fields("TempsShipping")

505       rstSoumCumulatif.Fields("NbrePersonne") = rstSoum.Fields("NbrePersonne")
510       rstSoumCumulatif.Fields("TempsHebergement") = rstSoum.Fields("TempsHebergement")
515       rstSoumCumulatif.Fields("TempsRepas") = rstSoum.Fields("TempsRepas")
520       rstSoumCumulatif.Fields("TempsTransport") = rstSoum.Fields("TempsTransport")
525       rstSoumCumulatif.Fields("TempsUniteMobile") = rstSoum.Fields("TempsUniteMobile")

530       rstSoumCumulatif.Fields("TotalHebergement") = rstSoum.Fields("TotalHebergement")
535       rstSoumCumulatif.Fields("TotalRepas") = rstSoum.Fields("TotalRepas")
540       rstSoumCumulatif.Fields("PrixEmballage") = rstSoum.Fields("PrixEmballage")

545       rstSoumCumulatif.Fields("total_manuel") = rstSoum.Fields("total_manuel")

550       rstSoumCumulatif.Fields("MontantForfait") = rstSoum.Fields("MontantForfait")
555     Else
560       Do While Not rstSoum.EOF
565         If Not IsNull(rstSoum.Fields("manuel")) Then
570           dblNbreManuel = dblNbreManuel + CDbl(rstSoum.Fields("manuel"))
575         End If

580         If Not IsNull(rstSoum.Fields("TempsMachinage")) Then
585           dblTempsMachinage = dblTempsMachinage + CDbl(rstSoum.Fields("TempsMachinage"))
590         End If

595         If Not IsNull(rstSoum.Fields("TempsCoupe")) Then
600           dblTempsCoupe = dblTempsCoupe + CDbl(rstSoum.Fields("TempsCoupe"))
605         End If

610         If Not IsNull(rstSoum.Fields("TempsSoudure")) Then
615           dblTempsSoudure = dblTempsSoudure + CDbl(rstSoum.Fields("TempsSoudure"))
620         End If

625         If Not IsNull(rstSoum.Fields("TempsAssemblage")) Then
630           dblTempsAssemblage = dblTempsAssemblage + CDbl(rstSoum.Fields("TempsAssemblage"))
635         End If

640         If Not IsNull(rstSoum.Fields("TempsPeinture")) Then
645           dblTempsPeinture = dblTempsPeinture + CDbl(rstSoum.Fields("TempsPeinture"))
650         End If

655         If Not IsNull(rstSoum.Fields("TempsTest")) Then
660           dblTempsTest = dblTempsTest + CDbl(rstSoum.Fields("TempsTest"))
665         End If

670         If Not IsNull(rstSoum.Fields("TempsDessin")) Then
675           dblTempsDessin = dblTempsDessin + CDbl(rstSoum.Fields("TempsDessin"))
680         End If

685         If Not IsNull(rstSoum.Fields("TempsFormation")) Then
690           dblTempsFormation = dblTempsFormation + CDbl(rstSoum.Fields("TempsFormation"))
695         End If

700         If Not IsNull(rstSoum.Fields("TempsInstallation")) Then
705           dblTempsInstallation = dblTempsInstallation + CDbl(rstSoum.Fields("TempsInstallation"))
710         End If

715         If Not IsNull(rstSoum.Fields("TempsGestion")) Then
720           dblTempsGestion = dblTempsGestion + CDbl(rstSoum.Fields("TempsGestion"))
725         End If

730         If Not IsNull(rstSoum.Fields("TempsShipping")) Then
735           dblTempsShipping = dblTempsShipping + CDbl(rstSoum.Fields("TempsShipping"))
740         End If

745         If Not IsNull(rstSoum.Fields("TempsTransport")) Then
750           dblTempsTransport = dblTempsTransport + CDbl(rstSoum.Fields("TempsTransport"))
755         End If

760         If Not IsNull(rstSoum.Fields("TempsUniteMobile")) Then
765           dblTempsUniteMobile = dblTempsUniteMobile + CDbl(rstSoum.Fields("TempsUniteMobile"))
770         End If

775         If Not IsNull(rstSoum.Fields("TotalHebergement")) Then
780           dblTotalHebergement = dblTotalHebergement + CDbl(rstSoum.Fields("TotalHebergement"))
785         End If

790         If Not IsNull(rstSoum.Fields("TotalRepas")) Then
795           dblTotalRepas = dblTotalRepas + CDbl(rstSoum.Fields("TotalRepas"))
800         End If

805         If Not IsNull(rstSoum.Fields("PrixEmballage")) Then
810           dblPrixEmballage = dblPrixEmballage + CDbl(rstSoum.Fields("PrixEmballage"))
815         End If

820         If Not IsNull(rstSoum.Fields("total_manuel")) Then
825           dblTotalManuel = dblTotalManuel + CDbl(rstSoum.Fields("total_manuel"))
830         End If

835         If Not IsNull(rstSoum.Fields("MontantForfait")) Then
840           If IsNumeric(rstSoum.Fields("MontantForfait")) Then
845             dblForfait = dblForfait + CDbl(rstSoum.Fields("MontantForfait"))
850           End If
855         End If

860         Call rstSoum.MoveNext
865       Loop

870       rstSoumCumulatif.Fields("manuel") = dblNbreManuel

875       rstSoumCumulatif.Fields("TempsMachinage") = dblTempsMachinage
880       rstSoumCumulatif.Fields("TempsCoupe") = dblTempsCoupe
885       rstSoumCumulatif.Fields("TempsSoudure") = dblTempsSoudure
890       rstSoumCumulatif.Fields("TempsAssemblage") = dblTempsAssemblage
895       rstSoumCumulatif.Fields("TempsPeinture") = dblTempsPeinture
900       rstSoumCumulatif.Fields("TempsTest") = dblTempsTest
905       rstSoumCumulatif.Fields("TempsDessin") = dblTempsDessin
910       rstSoumCumulatif.Fields("TempsFormation") = dblTempsFormation
915       rstSoumCumulatif.Fields("TempsInstallation") = dblTempsInstallation
920       rstSoumCumulatif.Fields("TempsGestion") = dblTempsGestion
925       rstSoumCumulatif.Fields("TempsShipping") = dblTempsShipping

930       rstSoumCumulatif.Fields("TempsTransport") = dblTempsTransport
935       rstSoumCumulatif.Fields("TempsUniteMobile") = dblTempsUniteMobile

940       rstSoumCumulatif.Fields("TotalHebergement") = dblTotalHebergement
945       rstSoumCumulatif.Fields("TotalRepas") = dblTotalRepas
950       rstSoumCumulatif.Fields("PrixEmballage") = dblPrixEmballage

955       rstSoumCumulatif.Fields("total_manuel") = dblTotalManuel

960       rstSoumCumulatif.Fields("MontantForfait") = dblForfait
965     End If

970     Call rstSoumCumulatif.Update

975     Call rstSoumCumulatif.Close

980     Call rstSoum.Close

        'AJOUT DES PIÈCES
985     Call rstPiecesCumulatif.Open("SELECT * FROM GRB_Soumission_Pieces WHERE IDSoumission = '" & sNoCumulatif & "'", g_connData, adOpenDynamic, adLockOptimistic)
             
990     If bCumulatifExiste = True Then
995       Call g_connData.Execute("DELETE * FROM GRB_Soumission_Pieces WHERE IDSoumission = '" & sNoCumulatif & "' AND Provenance = '" & Right$(txtNoProjSoum.Text, 2) & "'")

1000      Call rstPieces.Open("SELECT * FROM GRB_Soumission_Pieces WHERE IDSoumission = '" & txtNoProjSoum.Text & "' ORDER BY NuméroLigne", g_connData, adOpenForwardOnly, adLockReadOnly)
1005    Else
1010      Call rstPieces.Open("SELECT * FROM GRB_Soumission_Pieces WHERE LEFT(IDSoumission, 6) = '" & Left$(txtNoProjSoum.Text, 6) & "' AND RIGHT(IDSoumission, 2) <> '99' ORDER BY CInt(OrdreSection), NuméroLigne", g_connData, adOpenForwardOnly, adLockReadOnly)
1015    End If

1020    Do While Not rstPieces.EOF
1025      Call rstPiecesCumulatif.AddNew

1030      rstPiecesCumulatif.Fields("IDSoumission") = sNoCumulatif
1035      rstPiecesCumulatif.Fields("IDSection") = rstPieces.Fields("IDSection")
1040      rstPiecesCumulatif.Fields("NumItem") = rstPieces.Fields("NumItem")
1045      rstPiecesCumulatif.Fields("Qté") = rstPieces.Fields("Qté")
1050      rstPiecesCumulatif.Fields("Desc_FR") = rstPieces.Fields("Desc_FR")
1055      rstPiecesCumulatif.Fields("Desc_EN") = rstPieces.Fields("Desc_EN")
1060      rstPiecesCumulatif.Fields("Manufact") = rstPieces.Fields("Manufact")
1065      rstPiecesCumulatif.Fields("Prix_list") = rstPieces.Fields("Prix_list")
1070      rstPiecesCumulatif.Fields("Escompte") = rstPieces.Fields("Escompte")
1075      rstPiecesCumulatif.Fields("Prix_net") = rstPieces.Fields("Prix_net")
1080      rstPiecesCumulatif.Fields("IDFRS") = rstPieces.Fields("IDFRS")
1085      rstPiecesCumulatif.Fields("Prix_total") = rstPieces.Fields("Prix_total")
1090      rstPiecesCumulatif.Fields("Profit_Argent") = rstPieces.Fields("Profit_Argent")
1095      rstPiecesCumulatif.Fields("SousSection") = rstPieces.Fields("SousSection")
1100      rstPiecesCumulatif.Fields("OrdreSection") = rstPieces.Fields("OrdreSection")
1105      rstPiecesCumulatif.Fields("NuméroLigne") = rstPieces.Fields("NuméroLigne")
1110      rstPiecesCumulatif.Fields("PrixOrigine") = rstPieces.Fields("PrixOrigine")
1115      rstPiecesCumulatif.Fields("Type") = rstPieces.Fields("Type")
1120      rstPiecesCumulatif.Fields("Visible") = rstPieces.Fields("Visible")
1125      rstPiecesCumulatif.Fields("Commentaire") = rstPieces.Fields("Commentaire")
1130      rstPiecesCumulatif.Fields("Devise") = rstPieces.Fields("Devise")
1135      rstPiecesCumulatif.Fields("Provenance") = Right(rstPieces.Fields("IDSoumission"), 2)

1140      Call rstPiecesCumulatif.Update

1145      Call rstPieces.MoveNext
1150    Loop

1155    Call rstPiecesCumulatif.Close
1160    Call rstPieces.Close

1165    Set rstSoum = Nothing
1170    Set rstPieces = Nothing
1175    Set rstSoumCumulatif = Nothing
1180    Set rstPiecesCumulatif = Nothing

1185    Call CalculerTotalRecordset(sNoCumulatif)

1190    Exit Sub

AfficherErreur:

1195    woups "FrmProjSoumElec", "AjouterSoumissionAuCumulatif", Err, Erl
End Sub

Private Sub RecreerProjetCumulatif()

5       On Error GoTo AfficherErreur

10      Dim sNoCumulatif          As String
15      Dim rstProj               As ADODB.Recordset
20      Dim rstPieces             As ADODB.Recordset
25      Dim rstProjCumulatif      As ADODB.Recordset
30      Dim rstPiecesCumulatif    As ADODB.Recordset
35      Dim rstProjSoum           As ADODB.Recordset
40      Dim dblNbreManuel         As Double
45      Dim dblPrixEmballage      As Double
50      Dim dblTotalManuel        As Double
55      Dim dblForfait            As Double
60      Dim bSupprimer            As Boolean

65      sNoCumulatif = Left$(txtNoProjSoum.Text, 7) & "99"

70      Set rstProj = New ADODB.Recordset
75      Set rstPieces = New ADODB.Recordset
80      Set rstProjCumulatif = New ADODB.Recordset
85      Set rstPiecesCumulatif = New ADODB.Recordset

90      rstProj.CursorLocation = adUseClient

95      Call rstProj.Open("SELECT * FROM GRB_ProjetMec WHERE LEFT(IDProjet, 6) = '" & Left$(txtNoProjSoum.Text, 6) & "' AND RIGHT(IDProjet, 2) <> '99'", g_connData, adOpenForwardOnly, adLockReadOnly)

100     If rstProj.EOF Then
105       Call g_connData.Execute("DELETE * FROM GRB_ProjSoum WHERE IDProjSoum = '" & sNoCumulatif & "'")

110       Call g_connData.Execute("DELETE * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoCumulatif & "' AND Type = 'E'")

          'Efface le projet
115       Call g_connData.Execute("DELETE * FROM GRB_ProjetMec WHERE IDProjet = '" & sNoCumulatif & "'")

120       bSupprimer = True
125     Else
130       Call rstProjCumulatif.Open("SELECT * FROM GRB_ProjetMec WHERE IDProjet = '" & sNoCumulatif & "'", g_connData, adOpenDynamic, adLockOptimistic)

135       If rstProj.RecordCount = 1 Then
140         rstProjCumulatif.Fields("manuel") = rstProj.Fields("manuel")

145         rstProjCumulatif.Fields("PrixEmballage") = rstProj.Fields("PrixEmballage")

150         rstProjCumulatif.Fields("total_manuel") = rstProj.Fields("total_manuel")

155         rstProjCumulatif.Fields("MontantForfait") = rstProj.Fields("MontantForfait")
160       Else
165         Do While Not rstProj.EOF
170           If Not IsNull(rstProj.Fields("manuel")) Then
175             dblNbreManuel = dblNbreManuel + CDbl(rstProj.Fields("manuel"))
180           End If

185           If Not IsNull(rstProj.Fields("PrixEmballage")) Then
190             dblPrixEmballage = dblPrixEmballage + CDbl(rstProj.Fields("PrixEmballage"))
195           End If

200           If Not IsNull(rstProj.Fields("total_manuel")) Then
205             dblTotalManuel = dblTotalManuel + CDbl(rstProj.Fields("total_manuel"))
210           End If

215           If Not IsNull(rstProj.Fields("MontantForfait")) Then
220             If IsNumeric(rstProj.Fields("MontantForfait")) Then
225               dblForfait = dblForfait + CDbl(rstProj.Fields("MontantForfait"))
230             End If
235           End If

240           Call rstProj.MoveNext
245         Loop

250         rstProjCumulatif.Fields("manuel") = dblNbreManuel
255         rstProjCumulatif.Fields("PrixEmballage") = dblPrixEmballage
260         rstProjCumulatif.Fields("total_manuel") = dblTotalManuel
265         rstProjCumulatif.Fields("MontantForfait") = dblForfait
270       End If

275       Call rstProj.Close

280       Call rstProjCumulatif.Update

285       Call rstProjCumulatif.Close
290     End If

        'AJOUT DES PIÈCES
295     Call rstPiecesCumulatif.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoCumulatif & "'", g_connData, adOpenDynamic, adLockOptimistic)

300     Call g_connData.Execute("DELETE * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoCumulatif & "' AND Provenance = '" & Right$(txtNoProjSoum.Text, 2) & "'")

305     Call rstPieces.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & txtNoProjSoum.Text & "' AND Provenance Is Null OR Provenance = '' ORDER BY NuméroLigne", g_connData, adOpenForwardOnly, adLockReadOnly)

310     Do While Not rstPieces.EOF
315       Call rstPiecesCumulatif.AddNew

320       rstPiecesCumulatif.Fields("IDProjet") = sNoCumulatif
325       rstPiecesCumulatif.Fields("IDSection") = rstPieces.Fields("IDSection")
330       rstPiecesCumulatif.Fields("NumItem") = rstPieces.Fields("NumItem")
335       rstPiecesCumulatif.Fields("Qté") = rstPieces.Fields("Qté")
340       rstPiecesCumulatif.Fields("Desc_FR") = rstPieces.Fields("Desc_FR")
345       rstPiecesCumulatif.Fields("Desc_EN") = rstPieces.Fields("Desc_EN")
350       rstPiecesCumulatif.Fields("Manufact") = rstPieces.Fields("Manufact")
355       rstPiecesCumulatif.Fields("Prix_list") = rstPieces.Fields("Prix_list")
360       rstPiecesCumulatif.Fields("Escompte") = rstPieces.Fields("Escompte")
365       rstPiecesCumulatif.Fields("Prix_net") = rstPieces.Fields("Prix_net")
370       rstPiecesCumulatif.Fields("IDFRS") = rstPieces.Fields("IDFRS")
375       rstPiecesCumulatif.Fields("Prix_total") = rstPieces.Fields("Prix_total")
380       rstPiecesCumulatif.Fields("Profit_Argent") = rstPieces.Fields("Profit_Argent")
385       rstPiecesCumulatif.Fields("SousSection") = rstPieces.Fields("SousSection")
390       rstPiecesCumulatif.Fields("OrdreSection") = rstPieces.Fields("OrdreSection")
395       rstPiecesCumulatif.Fields("NuméroLigne") = rstPieces.Fields("NuméroLigne")
400       rstPiecesCumulatif.Fields("PrixOrigine") = rstPieces.Fields("PrixOrigine")
405       rstPiecesCumulatif.Fields("Type") = rstPieces.Fields("Type")
410       rstPiecesCumulatif.Fields("Visible") = rstPieces.Fields("Visible")
415       rstPiecesCumulatif.Fields("Commentaire") = rstPieces.Fields("Commentaire")
420       rstPiecesCumulatif.Fields("Devise") = rstPieces.Fields("Devise")
425       rstPiecesCumulatif.Fields("Provenance") = Right(rstPieces.Fields("IDProjet"), 2)

430       Call rstPiecesCumulatif.Update

435       Call rstPieces.MoveNext
440     Loop

445     Call rstPiecesCumulatif.Close
450     Call rstPieces.Close

455     Set rstProj = Nothing
460     Set rstPieces = Nothing
465     Set rstProjCumulatif = Nothing
470     Set rstPiecesCumulatif = Nothing

475     If bSupprimer = False Then
480       Call CalculerTotalRecordset(sNoCumulatif)
485     End If

490     Exit Sub

AfficherErreur:

495     woups "FrmProjSoumElec", "AjouterProjetAuCumulatif", Err, Erl
End Sub

Private Sub RecreerSoumissionCumulatif()

5       On Error GoTo AfficherErreur

10      Dim sNoCumulatif          As String
15      Dim rstSoum               As ADODB.Recordset
20      Dim rstPieces             As ADODB.Recordset
25      Dim rstSoumCumulatif      As ADODB.Recordset
30      Dim rstPiecesCumulatif    As ADODB.Recordset
35      Dim rstProjSoum           As ADODB.Recordset
40      Dim dblNbreManuel         As Double
45      Dim dblTempsMachinage     As Double
50      Dim dblTempsCoupe         As Double
55      Dim dblTempsSoudure       As Double
60      Dim dblTempsAssemblage    As Double
65      Dim dblTempsPeinture      As Double
70      Dim dblTempsTest          As Double
75      Dim dblTempsDessin        As Double
80      Dim dblTempsFormation     As Double
85      Dim dblTempsInstallation  As Double
90      Dim dblTempsGestion       As Double
95      Dim dblTempsShipping      As Double
100     Dim dblTempsTransport     As Double
105     Dim dblTempsUniteMobile   As Double
110     Dim dblTotalHebergement   As Double
115     Dim dblTotalRepas         As Double
120     Dim dblPrixEmballage      As Double
125     Dim dblTotalManuel        As Double
130     Dim dblForfait            As Double
135     Dim bSupprimer            As Boolean

140     sNoCumulatif = Left$(txtNoProjSoum.Text, 7) & "99"

145     Set rstSoum = New ADODB.Recordset
150     Set rstPieces = New ADODB.Recordset
155     Set rstSoumCumulatif = New ADODB.Recordset
160     Set rstPiecesCumulatif = New ADODB.Recordset
     
165     rstSoum.CursorLocation = adUseClient
     
170     Call rstSoum.Open("SELECT * FROM GRB_SoumissionMec WHERE LEFT(IDSoumission, 6) = '" & Left$(txtNoProjSoum.Text, 6) & "' AND RIGHT(IDSoumission, 2) <> '99'", g_connData, adOpenForwardOnly, adLockReadOnly)

175     If rstSoum.EOF Then
180       Call g_connData.Execute("DELETE * FROM GRB_ProjSoum WHERE IDProjSoum = '" & sNoCumulatif & "'")

185       Call g_connData.Execute("DELETE * FROM GRB_Soumission_Pieces WHERE IDSoumission = '" & sNoCumulatif & "' AND Type = 'E'")
                
          'Efface la soumission
190       Call g_connData.Execute("DELETE * FROM GRB_SoumissionMec WHERE IDSoumission = '" & sNoCumulatif & "'")

195       bSupprimer = True
200     Else
205       Call rstSoumCumulatif.Open("SELECT * FROM GRB_SoumissionMec WHERE IDSoumission = '" & sNoCumulatif & "'", g_connData, adOpenDynamic, adLockOptimistic)

210       If rstSoum.RecordCount = 1 Then
215         rstSoumCumulatif.Fields("manuel") = rstSoum.Fields("manuel")
  
220         rstSoumCumulatif.Fields("TempsMachinage") = rstSoum.Fields("TempsMachinage")
225         rstSoumCumulatif.Fields("TempsCoupe") = rstSoum.Fields("TempsCoupe")
230         rstSoumCumulatif.Fields("TempsSoudure") = rstSoum.Fields("TempsSoudure")
235         rstSoumCumulatif.Fields("TempsAssemblage") = rstSoum.Fields("TempsAssemblage")
240         rstSoumCumulatif.Fields("TempsPeinture") = rstSoum.Fields("TempsPeinture")
245         rstSoumCumulatif.Fields("TempsTest") = rstSoum.Fields("TempsTest")
250         rstSoumCumulatif.Fields("TempsDessin") = rstSoum.Fields("TempsDessin")
255         rstSoumCumulatif.Fields("TempsFormation") = rstSoum.Fields("TempsFormation")
260         rstSoumCumulatif.Fields("TempsInstallation") = rstSoum.Fields("TempsInstallation")
265         rstSoumCumulatif.Fields("TempsGestion") = rstSoum.Fields("TempsGestion")
270         rstSoumCumulatif.Fields("TempsShipping") = rstSoum.Fields("TempsShipping")
  
275         rstSoumCumulatif.Fields("NbrePersonne") = rstSoum.Fields("NbrePersonne")
280         rstSoumCumulatif.Fields("TempsHebergement") = rstSoum.Fields("TempsHebergement")
285         rstSoumCumulatif.Fields("TempsRepas") = rstSoum.Fields("TempsRepas")
290         rstSoumCumulatif.Fields("TempsTransport") = rstSoum.Fields("TempsTransport")
295         rstSoumCumulatif.Fields("TempsUniteMobile") = rstSoum.Fields("TempsUniteMobile")
  
300         rstSoumCumulatif.Fields("TotalHebergement") = rstSoum.Fields("TotalHebergement")
305         rstSoumCumulatif.Fields("TotalRepas") = rstSoum.Fields("TotalRepas")
310         rstSoumCumulatif.Fields("PrixEmballage") = rstSoum.Fields("PrixEmballage")

315         rstSoumCumulatif.Fields("total_manuel") = rstSoum.Fields("total_manuel")

320         rstSoumCumulatif.Fields("MontantForfait") = rstSoum.Fields("MontantForfait")
325       Else
330         Do While Not rstSoum.EOF
335           If Not IsNull(rstSoum.Fields("manuel")) Then
340             dblNbreManuel = dblNbreManuel + CDbl(rstSoum.Fields("manuel"))
345           End If
  
350           If Not IsNull(rstSoum.Fields("TempsMachinage")) Then
355             dblTempsMachinage = dblTempsMachinage + CDbl(rstSoum.Fields("TempsMachinage"))
360           End If
    
365           If Not IsNull(rstSoum.Fields("TempsCoupe")) Then
370             dblTempsCoupe = dblTempsCoupe + CDbl(rstSoum.Fields("TempsCoupe"))
375           End If
  
380           If Not IsNull(rstSoum.Fields("TempsSoudure")) Then
385             dblTempsSoudure = dblTempsSoudure + CDbl(rstSoum.Fields("TempsSoudure"))
390           End If
  
395           If Not IsNull(rstSoum.Fields("TempsAssemblage")) Then
400             dblTempsAssemblage = dblTempsAssemblage + CDbl(rstSoum.Fields("TempsAssemblage"))
405           End If
  
410           If Not IsNull(rstSoum.Fields("TempsPeinture")) Then
415             dblTempsPeinture = dblTempsPeinture + CDbl(rstSoum.Fields("TempsPeinture"))
420           End If
  
425           If Not IsNull(rstSoum.Fields("TempsTest")) Then
430             dblTempsTest = dblTempsTest + CDbl(rstSoum.Fields("TempsTest"))
435           End If
  
440           If Not IsNull(rstSoum.Fields("TempsDessin")) Then
445             dblTempsDessin = dblTempsDessin + CDbl(rstSoum.Fields("TempsDessin"))
450           End If
    
455           If Not IsNull(rstSoum.Fields("TempsFormation")) Then
460             dblTempsFormation = dblTempsFormation + CDbl(rstSoum.Fields("TempsFormation"))
465           End If
  
470           If Not IsNull(rstSoum.Fields("TempsInstallation")) Then
475             dblTempsInstallation = dblTempsInstallation + CDbl(rstSoum.Fields("TempsInstallation"))
480           End If
  
485           If Not IsNull(rstSoum.Fields("TempsGestion")) Then
490             dblTempsGestion = dblTempsGestion + CDbl(rstSoum.Fields("TempsGestion"))
495           End If

500           If Not IsNull(rstSoum.Fields("TempsShipping")) Then
505             dblTempsShipping = dblTempsShipping + CDbl(rstSoum.Fields("TempsShipping"))
510           End If
      
515           If Not IsNull(rstSoum.Fields("TempsTransport")) Then
520             dblTempsTransport = dblTempsTransport + CDbl(rstSoum.Fields("TempsTransport"))
525           End If
  
530           If Not IsNull(rstSoum.Fields("TempsUniteMobile")) Then
535             dblTempsUniteMobile = dblTempsUniteMobile + CDbl(rstSoum.Fields("TempsUniteMobile"))
540           End If
  
545           If Not IsNull(rstSoum.Fields("TotalHebergement")) Then
550             dblTotalHebergement = dblTotalHebergement + CDbl(rstSoum.Fields("TotalHebergement"))
555           End If
  
560           If Not IsNull(rstSoum.Fields("TotalRepas")) Then
565             dblTotalRepas = dblTotalRepas + CDbl(rstSoum.Fields("TotalRepas"))
570           End If
  
575           If Not IsNull(rstSoum.Fields("PrixEmballage")) Then
580             dblPrixEmballage = dblPrixEmballage + CDbl(rstSoum.Fields("PrixEmballage"))
585           End If
  
590           If Not IsNull(rstSoum.Fields("total_manuel")) Then
595             dblTotalManuel = dblTotalManuel + CDbl(rstSoum.Fields("total_manuel"))
600           End If

605           If Not IsNull(rstSoum.Fields("MontantForfait")) Then
610             If IsNumeric(rstSoum.Fields("MontantForfait")) Then
615               dblForfait = dblForfait + CDbl(rstSoum.Fields("MontantForfait"))
620             End If
625           End If
  
630           Call rstSoum.MoveNext
635         Loop
  
640         rstSoumCumulatif.Fields("manuel") = dblNbreManuel
  
645         rstSoumCumulatif.Fields("TempsMachinage") = dblTempsMachinage
650         rstSoumCumulatif.Fields("TempsCoupe") = dblTempsCoupe
655         rstSoumCumulatif.Fields("TempsSoudure") = dblTempsSoudure
660         rstSoumCumulatif.Fields("TempsAssemblage") = dblTempsAssemblage
665         rstSoumCumulatif.Fields("TempsPeinture") = dblTempsPeinture
670         rstSoumCumulatif.Fields("TempsTest") = dblTempsTest
675         rstSoumCumulatif.Fields("TempsDessin") = dblTempsDessin
680         rstSoumCumulatif.Fields("TempsFormation") = dblTempsFormation
685         rstSoumCumulatif.Fields("TempsInstallation") = dblTempsInstallation
690         rstSoumCumulatif.Fields("TempsGestion") = dblTempsGestion
695         rstSoumCumulatif.Fields("TempsShipping") = dblTempsShipping
  
700         rstSoumCumulatif.Fields("TempsTransport") = dblTempsTransport
705         rstSoumCumulatif.Fields("TempsUniteMobile") = dblTempsUniteMobile
  
710         rstSoumCumulatif.Fields("TotalHebergement") = dblTotalHebergement
715         rstSoumCumulatif.Fields("TotalRepas") = dblTotalRepas
720         rstSoumCumulatif.Fields("PrixEmballage") = dblPrixEmballage
  
725         rstSoumCumulatif.Fields("total_manuel") = dblTotalManuel

730         rstSoumCumulatif.Fields("MontantForfait") = dblForfait
735       End If

740       Call rstSoum.Close

745       Call rstSoumCumulatif.Update

750       Call rstSoumCumulatif.Close
755     End If

        'AJOUT DES PIÈCES
760     Call rstPiecesCumulatif.Open("SELECT * FROM GRB_Soumission_Pieces WHERE IDSoumission = '" & sNoCumulatif & "'", g_connData, adOpenDynamic, adLockOptimistic)
             
765     Call g_connData.Execute("DELETE * FROM GRB_Soumission_Pieces WHERE IDSoumission = '" & sNoCumulatif & "'")

770     Call rstPieces.Open("SELECT * FROM GRB_Soumission_Pieces WHERE LEFT(IDSoumission, 6) = '" & Left$(txtNoProjSoum.Text, 6) & "' AND RIGHT(IDSoumission, 2) <> '99' ORDER BY CInt(OrdreSection), NuméroLigne", g_connData, adOpenForwardOnly, adLockReadOnly)

775     Do While Not rstPieces.EOF
780       Call rstPiecesCumulatif.AddNew

785       rstPiecesCumulatif.Fields("IDSoumission") = sNoCumulatif
790       rstPiecesCumulatif.Fields("IDSection") = rstPieces.Fields("IDSection")
795       rstPiecesCumulatif.Fields("NumItem") = rstPieces.Fields("NumItem")
800       rstPiecesCumulatif.Fields("Qté") = rstPieces.Fields("Qté")
805       rstPiecesCumulatif.Fields("Desc_FR") = rstPieces.Fields("Desc_FR")
810       rstPiecesCumulatif.Fields("Desc_EN") = rstPieces.Fields("Desc_EN")
815       rstPiecesCumulatif.Fields("Manufact") = rstPieces.Fields("Manufact")
820       rstPiecesCumulatif.Fields("Prix_list") = rstPieces.Fields("Prix_list")
825       rstPiecesCumulatif.Fields("Escompte") = rstPieces.Fields("Escompte")
830       rstPiecesCumulatif.Fields("Prix_net") = rstPieces.Fields("Prix_net")
835       rstPiecesCumulatif.Fields("IDFRS") = rstPieces.Fields("IDFRS")
840       rstPiecesCumulatif.Fields("Prix_total") = rstPieces.Fields("Prix_total")
845       rstPiecesCumulatif.Fields("Profit_Argent") = rstPieces.Fields("Profit_Argent")
850       rstPiecesCumulatif.Fields("SousSection") = rstPieces.Fields("SousSection")
855       rstPiecesCumulatif.Fields("OrdreSection") = rstPieces.Fields("OrdreSection")
860       rstPiecesCumulatif.Fields("NuméroLigne") = rstPieces.Fields("NuméroLigne")
865       rstPiecesCumulatif.Fields("PrixOrigine") = rstPieces.Fields("PrixOrigine")
870       rstPiecesCumulatif.Fields("Type") = rstPieces.Fields("Type")
875       rstPiecesCumulatif.Fields("Visible") = rstPieces.Fields("Visible")
880       rstPiecesCumulatif.Fields("Commentaire") = rstPieces.Fields("Commentaire")
885       rstPiecesCumulatif.Fields("Devise") = rstPieces.Fields("Devise")
890       rstPiecesCumulatif.Fields("Provenance") = Right(rstPieces.Fields("IDSoumission"), 2)

895       Call rstPiecesCumulatif.Update

900       Call rstPieces.MoveNext
905     Loop

910     Call rstPiecesCumulatif.Close
915     Call rstPieces.Close

920     Set rstSoum = New ADODB.Recordset
925     Set rstPieces = New ADODB.Recordset
930     Set rstSoumCumulatif = New ADODB.Recordset
935     Set rstPiecesCumulatif = New ADODB.Recordset

940     If bSupprimer = False Then
945       Call CalculerTotalRecordset(sNoCumulatif)
950     End If

955     Exit Sub

AfficherErreur:

960     woups "FrmProjSoumElec", "RecreerSoumissionCumulatif", Err, Erl
End Sub
Private Function ExportdansExcel(ByVal oRecordset As ADODB.Recordset)

5       On Error GoTo AfficherErreur

6       Dim iCount As Integer
10      Dim oXLApp As Excel.Application         'Declare the object variables
15      Dim oXLBook As Excel.Workbook
20      Dim oXLSheet As Excel.Worksheet

25      Set oXLApp = New Excel.Application    'Create a new instance of Excel
30      Set oXLBook = oXLApp.Workbooks.Add    'Add a new workbook
35      Set oXLSheet = oXLBook.Worksheets(1)  'Work with the first worksheet


' ajustement largeur des colonne
oXLSheet.Columns(1).ColumnWidth = 10
oXLSheet.Columns(2).ColumnWidth = 8
oXLSheet.Columns(3).ColumnWidth = 20
oXLSheet.Columns(4).ColumnWidth = 45
oXLSheet.Columns(5).ColumnWidth = 20
oXLSheet.Columns(6).ColumnWidth = 12
oXLSheet.Columns(7).ColumnWidth = 12
oXLSheet.Columns(8).ColumnWidth = 12
oXLSheet.Columns(9).ColumnWidth = 12
oXLSheet.Columns(10).ColumnWidth = 30
oXLSheet.Columns(11).ColumnWidth = 20
oXLSheet.Columns(12).ColumnWidth = 20

oXLSheet.Range("A1: N1").Font.Bold = True




40      With oXLSheet                    'Fill with data
45        For iCount = 0 To (oRecordset.Fields.count - 1)
50          .Cells(1, iCount + 1) = oRecordset.Fields(iCount).Name
55        Next iCount
60          'create and fill a recordset here, called oRecordset
65        .Range("A2").CopyFromRecordset oRecordset
70      End With

75      oXLApp.Visible = True                'Show it to the user
80      Set oXLSheet = Nothing               'Disconnect from all Excel objects (let the user take over)
85      Set oXLBook = Nothing
90      Set oXLApp = Nothing

95      Exit Function

AfficherErreur:

960     woups "FrmProjSoumMec", "ExportdansExcel", Err, Erl
End Function
