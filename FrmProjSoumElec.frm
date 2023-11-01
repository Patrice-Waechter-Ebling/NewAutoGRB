VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmProjSoumElec 
   BackColor       =   &H00000000&
   Caption         =   "Projets / Soumissions Électriques"
   ClientHeight    =   8265
   ClientLeft      =   225
   ClientTop       =   645
   ClientWidth     =   13380
   Icon            =   "FrmProjSoumElec.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "FrmProjSoumElec.frx":2CFA
   ScaleHeight     =   8265
   ScaleWidth      =   13380
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbOuvertFerme 
      Height          =   315
      ItemData        =   "FrmProjSoumElec.frx":5C07
      Left            =   4560
      List            =   "FrmProjSoumElec.frx":5C11
      Style           =   2  'Dropdown List
      TabIndex        =   140
      Top             =   120
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog dlgExcel 
      Left            =   120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   360
      TabIndex        =   20
      Top             =   600
      Visible         =   0   'False
      Width           =   11415
      Begin VB.CommandButton cmdOKFRS 
         Caption         =   "OK"
         Height          =   375
         Left            =   10200
         TabIndex        =   24
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdAnnulerFRS 
         Caption         =   "Annuler"
         Height          =   375
         Left            =   9000
         TabIndex        =   23
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdSupprimerFRS 
         Caption         =   "Supprimer"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   1920
         Width           =   975
      End
      Begin MSComctlLib.ListView lvwFournisseur 
         Height          =   1575
         Left            =   120
         TabIndex        =   21
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
      Left            =   360
      TabIndex        =   15
      Top             =   600
      Visible         =   0   'False
      Width           =   10335
      Begin VB.CommandButton cmdAnnulerPieceTrouve 
         Caption         =   "Annuler"
         Height          =   375
         Left            =   7920
         TabIndex        =   17
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdOKPieceTrouve 
         Caption         =   "OK"
         Height          =   375
         Left            =   9120
         TabIndex        =   18
         Top             =   2280
         Width           =   1095
      End
      Begin MSComctlLib.ListView lvwPieceTrouve 
         Height          =   1935
         Left            =   120
         TabIndex        =   16
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
   Begin MSComctlLib.ListView lvwHistorique 
      Height          =   1575
      Left            =   120
      TabIndex        =   55
      Top             =   1320
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
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
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvwBavard 
      Height          =   1575
      Left            =   120
      TabIndex        =   54
      Top             =   1320
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
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
   Begin MSComCtl2.MonthView mvwDateFacturation 
      Height          =   2370
      Left            =   720
      TabIndex        =   19
      Top             =   600
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   0
      Appearance      =   1
      StartOfWeek     =   106364929
      CurrentDate     =   37761
   End
   Begin MSComCtl2.MonthView mvwDate 
      Height          =   2370
      Left            =   9360
      TabIndex        =   41
      Top             =   720
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   0
      Appearance      =   1
      StartOfWeek     =   106364929
      CurrentDate     =   37761
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
      Left            =   3600
      TabIndex        =   89
      Top             =   4560
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton cmdAnnulerDateRequise 
         Caption         =   "Annuler"
         Height          =   375
         Left            =   3480
         TabIndex        =   95
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdOKDateRequise 
         Caption         =   "OK"
         Height          =   375
         Left            =   3480
         TabIndex        =   93
         Top             =   720
         Width           =   1095
      End
      Begin MSComCtl2.MonthView mvwDateRequise 
         Height          =   2370
         Left            =   600
         TabIndex        =   91
         Top             =   360
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   0
         Appearance      =   1
         StartOfWeek     =   106364929
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
      Left            =   3600
      TabIndex        =   96
      Top             =   4560
      Visible         =   0   'False
      Width           =   4815
      Begin VB.TextBox txtCommentaire 
         Height          =   2415
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   97
         Top             =   360
         Width           =   3255
      End
      Begin VB.CommandButton cmdOKCommentaire 
         Caption         =   "OK"
         Height          =   375
         Left            =   3600
         TabIndex        =   98
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdAnnulerCommentaire 
         Caption         =   "Annuler"
         Height          =   375
         Left            =   3600
         TabIndex        =   99
         Top             =   1200
         Width           =   1095
      End
   End
   Begin VB.Frame fraPrixPiece 
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
      Left            =   840
      TabIndex        =   76
      Top             =   4560
      Visible         =   0   'False
      Width           =   8895
      Begin VB.TextBox txtPrixSpecial 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4920
         TabIndex        =   86
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtPrixList 
         BackColor       =   &H00FFFFFF&
         DataField       =   "PRIX_LIST"
         DataSource      =   "DatCat1"
         Height          =   285
         Left            =   4920
         TabIndex        =   80
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtPrixNet 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4920
         TabIndex        =   84
         Top             =   1200
         Width           =   855
      End
      Begin VB.ComboBox cmbfrs 
         Height          =   315
         Left            =   240
         TabIndex        =   78
         Top             =   480
         Width           =   2775
      End
      Begin VB.OptionButton optUSA 
         BackColor       =   &H00000000&
         Caption         =   "USA"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7320
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
         TabIndex        =   87
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optSpain 
         BackColor       =   &H00000000&
         Caption         =   "SPA"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   8040
         TabIndex        =   90
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdOKPrix 
         Caption         =   "OK"
         Height          =   375
         Left            =   7440
         TabIndex        =   94
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdAnnulerPrix 
         Caption         =   "Annuler"
         Height          =   375
         Left            =   6240
         TabIndex        =   92
         Top             =   1800
         Width           =   1095
      End
      Begin MSMask.MaskEdBox mskEscompte 
         Height          =   255
         Left            =   4920
         TabIndex        =   82
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
         TabIndex        =   85
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Image imgCanada 
         Height          =   1065
         Left            =   6840
         Picture         =   "FrmProjSoumElec.frx":5C26
         Stretch         =   -1  'True
         Top             =   600
         Visible         =   0   'False
         Width           =   1680
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
         TabIndex        =   77
         Top             =   240
         Width           =   1215
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
         TabIndex        =   79
         Top             =   480
         Width           =   975
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
         TabIndex        =   81
         Top             =   840
         Width           =   1095
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
         TabIndex        =   83
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Image imgSpain 
         Height          =   1065
         Left            =   6840
         Picture         =   "FrmProjSoumElec.frx":5BC08
         Stretch         =   -1  'True
         Top             =   600
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Image imgEU 
         Height          =   1065
         Left            =   6840
         Picture         =   "FrmProjSoumElec.frx":5E097
         Stretch         =   -1  'True
         Top             =   600
         Visible         =   0   'False
         Width           =   1680
      End
   End
   Begin VB.CommandButton cmdAnglaisFrancais 
      Caption         =   "Anglais"
      Height          =   375
      Left            =   1200
      TabIndex        =   102
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   3360
      TabIndex        =   107
      Top             =   7320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdRapportFACT 
      Caption         =   "Fact"
      Height          =   375
      Left            =   2280
      TabIndex        =   106
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdTexte 
      Caption         =   "Texte"
      Height          =   375
      Left            =   120
      TabIndex        =   101
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdDemande 
      Caption         =   "Demande"
      Height          =   375
      Left            =   2280
      TabIndex        =   103
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdSortieMagasin 
      Caption         =   "Magasin"
      Height          =   375
      Left            =   5400
      TabIndex        =   111
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdMauvaisPrix 
      Caption         =   "Prix"
      Height          =   375
      Left            =   3240
      TabIndex        =   105
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdMaterielInutile 
      Caption         =   "Inutile"
      Height          =   375
      Left            =   5400
      TabIndex        =   113
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdCreerProjet 
      Caption         =   "Créer proj."
      Height          =   375
      Left            =   4320
      TabIndex        =   109
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdCatalogue 
      Caption         =   "Catalogue"
      Height          =   375
      Left            =   5400
      TabIndex        =   112
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdBonCommande 
      Caption         =   "Bon Com."
      Height          =   375
      Left            =   3360
      TabIndex        =   110
      Top             =   7320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdRetour 
      Caption         =   "Retour"
      Height          =   375
      Left            =   5400
      TabIndex        =   114
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdExtra 
      Caption         =   "Extra"
      Height          =   375
      Left            =   2280
      TabIndex        =   104
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdCopier 
      Caption         =   "Copier"
      Height          =   375
      Left            =   6480
      TabIndex        =   115
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdImprimer 
      Caption         =   "Imprimer"
      Height          =   375
      Left            =   120
      TabIndex        =   100
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdAjouter 
      Caption         =   "Ajouter"
      Height          =   375
      Left            =   7560
      TabIndex        =   116
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdFermer 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   10800
      TabIndex        =   121
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdAnnuler 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   9720
      TabIndex        =   119
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdEnregistrer 
      Caption         =   "Enregistrer"
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
      TabIndex        =   120
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdSupprimer 
      Caption         =   "Supprimer"
      Height          =   375
      Left            =   8640
      TabIndex        =   118
      Top             =   7320
      Width           =   975
   End
   Begin VB.Frame fraCertifDelais 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1095
      Left            =   9600
      TabIndex        =   42
      Top             =   1800
      Width           =   2175
      Begin VB.PictureBox picApprob 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   0
         ScaleHeight     =   495
         ScaleWidth      =   2295
         TabIndex        =   43
         Top             =   120
         Width           =   2295
         Begin VB.CheckBox chkUR 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            Caption         =   "UR"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   1560
            TabIndex        =   46
            Top             =   0
            Width           =   615
         End
         Begin VB.CheckBox chkCE 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            Caption         =   "CE"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   1560
            TabIndex        =   49
            Top             =   240
            Width           =   615
         End
         Begin VB.CheckBox chkCUR 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            Caption         =   "cUR"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   720
            TabIndex        =   47
            Top             =   240
            Width           =   735
         End
         Begin VB.CheckBox chkUL 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            Caption         =   "UL"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   840
            TabIndex        =   45
            Top             =   0
            Width           =   615
         End
         Begin VB.CheckBox chkCUL 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            Caption         =   "cUL"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   48
            Top             =   240
            Width           =   735
         End
         Begin VB.CheckBox chkCSA 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            Caption         =   "CSA"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   44
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.TextBox txtDelais 
         Height          =   288
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   780
         Width           =   975
      End
      Begin VB.CommandButton cmdDate 
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
         Left            =   1800
         TabIndex        =   52
         Top             =   780
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Delais"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   51
         Top             =   780
         Width           =   495
      End
   End
   Begin VB.Frame fraFsTransMarq 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   9960
      TabIndex        =   33
      Top             =   120
      Width           =   2055
      Begin VB.TextBox txtPrixSoumission 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   840
         Width           =   975
      End
      Begin VB.ComboBox cmbTransport 
         Height          =   315
         ItemData        =   "FrmProjSoumElec.frx":AAE09
         Left            =   960
         List            =   "FrmProjSoumElec.frx":AAE13
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtTransport 
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   0
         Width           =   1095
      End
      Begin VB.TextBox txtPrixReception 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblPrixSoumission 
         BackStyle       =   0  'Transparent
         Caption         =   "$ Soumission : "
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   39
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblPrixReception 
         BackStyle       =   0  'Transparent
         Caption         =   "$ Réception : "
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   37
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Transport"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdRafraichir 
      Caption         =   "Rafraichir"
      Height          =   315
      Left            =   8280
      TabIndex        =   66
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdTri 
      Caption         =   "Trier"
      Height          =   315
      Left            =   8280
      TabIndex        =   73
      Top             =   2700
      Width           =   975
   End
   Begin VB.Frame fraPrix 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2415
      Left            =   12360
      TabIndex        =   53
      Top             =   120
      Width           =   2655
      Begin VB.TextBox txtTotalTemps 
         Alignment       =   1  'Right Justify
         Height          =   288
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   128
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox txtTotalPieces 
         Alignment       =   1  'Right Justify
         Height          =   288
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   127
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtProfit 
         Alignment       =   1  'Right Justify
         Height          =   288
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   126
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtImprevus 
         Alignment       =   1  'Right Justify
         Height          =   288
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   125
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtCommission 
         Alignment       =   1  'Right Justify
         Height          =   288
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   124
         Top             =   1560
         Width           =   1335
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
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   123
         Text            =   "0"
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Administration"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   134
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label7 
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
         Left            =   240
         TabIndex        =   133
         Top             =   1920
         Width           =   885
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Profit"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   132
         Top             =   840
         Width           =   885
      End
      Begin VB.Label lblTotalPieces 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Pièces"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   131
         Top             =   480
         Width           =   885
      End
      Begin VB.Label lblImprevus 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Imprévus"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   130
         Top             =   1200
         Width           =   885
      End
      Begin VB.Label lblTotalTemps 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Temps"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   129
         Top             =   120
         Width           =   885
      End
   End
   Begin VB.ComboBox cmbTri 
      Height          =   315
      ItemData        =   "FrmProjSoumElec.frx":AAE2B
      Left            =   6480
      List            =   "FrmProjSoumElec.frx":AAE3E
      Style           =   2  'Dropdown List
      TabIndex        =   72
      Top             =   2640
      Width           =   1695
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
      ItemData        =   "FrmProjSoumElec.frx":AAE95
      Left            =   4800
      List            =   "FrmProjSoumElec.frx":AAE97
      TabIndex        =   71
      Top             =   2640
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtCheminPhotos 
      Height          =   285
      Left            =   720
      TabIndex        =   25
      Top             =   780
      Width           =   2295
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   255
      Left            =   3120
      TabIndex        =   27
      Top             =   780
      Width           =   255
   End
   Begin VB.CommandButton cmdPhotos 
      Caption         =   "Afficher"
      Height          =   255
      Left            =   3480
      TabIndex        =   28
      Top             =   780
      Width           =   855
   End
   Begin VB.TextBox txtProjet 
      Height          =   285
      Left            =   6000
      TabIndex        =   11
      Top             =   1320
      Width           =   3735
   End
   Begin VB.ComboBox cmbContact 
      Height          =   315
      Left            =   6000
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   960
      Width           =   2655
   End
   Begin VB.TextBox txtContact 
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   960
      Width           =   2655
   End
   Begin VB.ComboBox cmbClient 
      Height          =   315
      Left            =   6000
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   600
      Width           =   2655
   End
   Begin VB.TextBox txtClient 
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   600
      Width           =   2655
   End
   Begin VB.TextBox txtNoSoumission 
      ForeColor       =   &H00808080&
      Height          =   288
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.ComboBox cmbProjSoum 
      Height          =   315
      Left            =   6000
      TabIndex        =   3
      Text            =   "cmbProjSoum"
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtNoProjSoum 
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox cmbChoix 
      Height          =   315
      ItemData        =   "FrmProjSoumElec.frx":AAE99
      Left            =   3240
      List            =   "FrmProjSoumElec.frx":AAEA3
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtChoix 
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Timer tmrTemps 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   600
   End
   Begin VB.CommandButton cmdHistorique 
      Caption         =   "Historique des modifications"
      Height          =   495
      Left            =   120
      TabIndex        =   59
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdLegende 
      Caption         =   "Légende"
      Height          =   375
      Left            =   1680
      TabIndex        =   60
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdBavards 
      Caption         =   "Bavard"
      Height          =   375
      Left            =   2640
      TabIndex        =   61
      Top             =   2160
      Width           =   855
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
      TabIndex        =   69
      Top             =   2640
      Visible         =   0   'False
      Width           =   375
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
      ItemData        =   "FrmProjSoumElec.frx":AAEBB
      Left            =   1080
      List            =   "FrmProjSoumElec.frx":AAEC2
      Style           =   2  'Dropdown List
      TabIndex        =   67
      Top             =   2640
      Visible         =   0   'False
      Width           =   2175
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
      TabIndex        =   64
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtDateFacturation 
      Height          =   288
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   63
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdTemps 
      Caption         =   "Temps"
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox txtForfait 
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdForfait 
      Caption         =   "..."
      Height          =   285
      Left            =   1680
      TabIndex        =   56
      ToolTipText     =   "Ajoute un forfait"
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton cmdEffacerForfait 
      Caption         =   "Effacer"
      Height          =   285
      Left            =   2160
      TabIndex        =   57
      ToolTipText     =   "Efface le forfait"
      Top             =   1560
      Width           =   855
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
      TabIndex        =   75
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
      Left            =   3600
      TabIndex        =   62
      Top             =   2160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdReception 
      Caption         =   "Réception"
      Height          =   375
      Left            =   3240
      TabIndex        =   108
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdRechercherClient 
      Caption         =   "..."
      Height          =   315
      Left            =   8760
      TabIndex        =   122
      Top             =   600
      Width           =   375
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
      Left            =   5160
      TabIndex        =   135
      Top             =   1800
      Width           =   1935
      Begin VB.TextBox txtNbreManuel 
         Height          =   288
         Left            =   480
         MaxLength       =   4
         TabIndex        =   137
         Text            =   "0"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtPrixManuel 
         Height          =   288
         Left            =   1320
         TabIndex        =   136
         Text            =   "0"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Nbre"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   139
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Prix"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   138
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Label lblForfaitInitiale 
      BackStyle       =   0  'Transparent
      Caption         =   "Par : "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   32
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Forfait :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   30
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblDateFacturation 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date facturation"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   58
      Top             =   2040
      Width           =   1140
   End
   Begin VB.Label lblPasTemps 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Vérifier le temps"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Photos : "
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   26
      Top             =   780
      Width           =   630
   End
   Begin VB.Label lblNoSoumission 
      BackStyle       =   0  'Transparent
      Caption         =   "Soumission"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7800
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
      TabIndex        =   12
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
      TabIndex        =   7
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
      TabIndex        =   68
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
      TabIndex        =   70
      Top             =   2640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblProjet 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5040
      TabIndex        =   10
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblTri 
      BackStyle       =   0  'Transparent
      Caption         =   "Trier par :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6480
      TabIndex        =   65
      Top             =   2400
      Width           =   855
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
      Begin VB.Menu mnuID 
         Caption         =   "Ajouter / Modifier l'ID"
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
Attribute VB_Name = "FrmProjSoumElec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'**************************************************************************************'
'   Liste des valeurs du .Text et du .Tag pour chacune des colonnes de lvwSoumission   '
'**************************************************************************************'
'        COLONNE            |           TEXTE            |          TAG                '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_QUANTITE        |   Quantité                 |     ID Section              '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_PIECE           |   Pièce                    |     Sous-section            '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_DESCR           |   Description FR ou EN     |     Description FR ou EN    '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_MANUFACT        |   Manufacturier            |     Ordre section           '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_PRIX_LIST       |   Prix listé               |     Prix d'origine          '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_ESCOMPTE        |   Escompte                 |                             '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_PRIX_NET        |   Prix net                 |     Date de réception       '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_DISTRIB         |   Fournisseur              |     ID Fournisseur          '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_TEMPS           |   Temps                    |                             '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_MONTAGE         |   Montage                  |                             '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_TOTAL           |   Total                    |     Devise monétaire        '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_PROFIT          |   Profit                   | EXTRA, RETOUR ou ANNULATION '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_COMMENTAIRE     |   Commentaire              |                             '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_ID              |   ID                       |                             '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_FACTURATION     |   Facturation              |                             '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_DATE_COMMANDE   |   Date Commande            |     Numéro Retour           '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_DATE_REQUISE    |   Date Requise             |     Date Retour             '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_NOM_COMMANDE    |   Personne qui a commandé  |                             '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_NO_SEQUENTIEL   |   Numéro séquentiel du BC  |                             '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_PROVENANCE      |   Provenance               |                             '
'---------------------------|----------------------------|-----------------------------'

'Index des colonnes de lvwSoumission
Private Const I_COL_SOUM_QUANTITE         As Integer = 0
Private Const I_COL_SOUM_PIECE            As Integer = 1
Private Const I_COL_SOUM_DESCR            As Integer = 2
Private Const I_COL_SOUM_MANUFACT         As Integer = 3
Private Const I_COL_SOUM_PRIX_LIST        As Integer = 4
Private Const I_COL_SOUM_ESCOMPTE         As Integer = 5
Private Const I_COL_SOUM_PRIX_NET         As Integer = 6
Private Const I_COL_SOUM_DISTRIB          As Integer = 7
Private Const I_COL_SOUM_TEMPS            As Integer = 8
Private Const I_COL_SOUM_MONTAGE          As Integer = 9
Private Const I_COL_SOUM_TOTAL            As Integer = 10
Private Const I_COL_SOUM_PROFIT           As Integer = 11
Private Const I_COL_SOUM_COMMENTAIRE      As Integer = 12
Private Const I_COL_SOUM_ID               As Integer = 13
Private Const I_COL_SOUM_FACTURATION      As Integer = 14
Private Const I_COL_SOUM_DATE_COMMANDE    As Integer = 15
Private Const I_COL_SOUM_DATE_REQUISE     As Integer = 16
Private Const I_COL_SOUM_NOM_COMMANDE     As Integer = 17
Private Const I_COL_SOUM_NO_SEQUENTIEL    As Integer = 18
Private Const I_COL_SOUM_PROVENANCE       As Integer = 19

Private Const I_COL_SOUMISSION_PROV       As Integer = 13

'Index des colonnes de lvwSoumission si les colonnes contenant
'des prix ne sont pas là. (SP est pour Sans Prix)
Private Const I_COL_SOUM_SP_QUANTITE      As Integer = 0
Private Const I_COL_SOUM_SP_PIECE         As Integer = 1
Private Const I_COL_SOUM_SP_DESCR         As Integer = 2
Private Const I_COL_SOUM_SP_MANUFACT      As Integer = 3
Private Const I_COL_SOUM_SP_DISTRIB       As Integer = 4
Private Const I_COL_SOUM_SP_TEMPS         As Integer = 5
Private Const I_COL_SOUM_SP_MONTAGE       As Integer = 6
Private Const I_COL_SOUM_SP_COMMENTAIRE   As Integer = 7
Private Const I_COL_SOUM_SP_ID            As Integer = 8
Private Const I_COL_SOUM_SP_DATE_COMMANDE As Integer = 9
Private Const I_COL_SOUM_SP_DATE_REQUISE  As Integer = 10
Private Const I_COL_SOUM_SP_NOM_COMMANDE  As Integer = 11
Private Const I_COL_SOUM_SP_NO_SEQUENTIEL As Integer = 12
Private Const I_COL_SOUM_SP_PROVENANCE    As Integer = 13

Private Const I_COL_SOUMISSION_SP_PROV    As Integer = 8

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

'Index des transports
Private Const I_TRANS_FAB_GRANBY          As Integer = 0
Private Const I_TRANS_CLIENT              As Integer = 1

'Index de m_collFloorStock
Private Const I_IDX_FS_DIX_MOINS          As Integer = 1
Private Const I_IDX_FS_DIX                As Integer = 2
Private Const I_IDX_FS_QUINZE             As Integer = 3
Private Const I_IDX_FS_VINGT              As Integer = 4
Private Const I_IDX_FS_VINGT_CINQ         As Integer = 5
Private Const I_IDX_FS_CINQUANTE          As Integer = 6
Private Const I_IDX_FS_CENT               As Integer = 7

'Index de cmbChoix
Private Const I_IDX_SOUMISSION            As Integer = 0
Private Const I_IDX_PROJET                As Integer = 1

'Index de cmbOuvertFerme
Private Const I_CMB_OUVERT                As Integer = 0
Private Const I_CMB_TOUS                  As Integer = 1

'Index de cmbTri
Private Const I_CMB_PIECE_GRB             As Integer = 0
Private Const I_CMB_PIECE                 As Integer = 1
Private Const I_CMB_FABRICANT             As Integer = 2
Private Const I_CMB_DESCR_FR              As Integer = 3
Private Const I_CMB_DESCR_EN              As Integer = 4

'Constante s'il n'y a pas de sous-sections
Private Const S_PAS_SOUS_SECTION          As String = "PAS DE SOUS-SECTION"

'Valeur servant au resize du lvwSoumission si le form est agrandi
Private Const I_TOP_AFFICHAGE             As Integer = 3000
Private Const I_HEIGHT_AFFICHAGE          As Integer = 3930

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
  sTemps       As String
  sMontage     As String
  sTotal       As String
  sProfit      As String
  sDescrTag    As String
  sPrixListTag As String
  sFRSTag      As String
  lColor       As Long
End Type

'Variables pour la configurations
Private m_sProfit                As String
Private m_sCommission            As String
Private m_sImprevue              As String

'Pour la recherche de pièce dans lvwPieces
Private m_sTri                   As String

'Pour savoir quelle colonne trier
Private m_iCol                   As Integer

'Pour savoir si le form a déjà été sur l'événement resize
Private m_bResize                As Boolean

'Modes du form
Private m_bModeAjout             As Boolean
Private m_bModeAffichage         As Boolean

'Pour avoir une sous-section par défaut
Private m_sSousSection           As String

'Pour savoir si le form affiche les projets ou les soumissions
Private m_eType                  As enumType

'Pour savoir si les prix sont cachés ou non
Public m_bDroitPrix              As Boolean

'Pour ne pas être obligé d'ouvrir le recordset à chaque fois
Private m_bModifProj             As Boolean
Private m_bModifSoum             As Boolean
Private m_bModifBonCommande      As Boolean

'Variable pour savoir si l'utilisateur a le droit de voir le combo ou non
Private m_bComboChoix            As Boolean

Private m_eMode                  As enumMode

Private m_eLangage               As enumLangage

'Pour faire afficher le dernier enregitrement visionné après un ajout ou une
'modification
Private m_sAncienProjSoum        As String

Private m_bSupprimer             As Boolean

'Pour savoir si il faut calculer le temps mécanique ou non
Public m_bSansTemps              As Boolean

'Pour savoir si le lvwFournisseur est affiché après marchandise non utilisée
Private m_bPieceInutile          As Boolean

Public m_bAnnulerChemin          As Boolean
Public m_sChemin                 As String

Private m_bRecherchePiece        As Boolean

'Pour savoir si le changement de prix a été appelé à partir du bouton "Mauvais Prix"
Private m_bMauvaisPrix           As Boolean
Private m_bEnregistrement        As Boolean
Private m_collDateSupp           As Collection
Private m_collHeureSupp          As Collection
Private m_collQteSupp            As Collection
Private m_collNoItemSupp         As Collection
Private m_bChangementFRS         As Boolean

Public m_sTempsDessin            As String
Public m_sTempsFabrication       As String
Public m_sTempsAssemblage        As String
Public m_sTempsProgInterface     As String
Public m_sTempsProgAutomate      As String
Public m_sTempsProgRobot         As String
Public m_sTempsVision            As String
Public m_sTempsTest              As String
Public m_sTempsInstallation      As String
Public m_sTempsMiseService       As String
Public m_sTempsFormation         As String
Public m_sTempsGestion           As String
Public m_sTempsShipping          As String

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
Public m_sTauxFabrication        As String
Public m_sTauxAssemblage         As String
Public m_sTauxProgInterface      As String
Public m_sTauxProgAutomate       As String
Public m_sTauxProgRobot          As String
Public m_sTauxVision             As String
Public m_sTauxTest               As String
Public m_sTauxInstallation       As String
Public m_sTauxMiseService        As String
Public m_sTauxFormation          As String
Public m_sTauxGestion            As String
Public m_sTauxShipping           As String

Public m_bTempsDejaOuvert        As Boolean

Private m_sTexteRecherche        As String
Private m_arr_tyCopie()          As tyCopiePiece
Private m_iNbreCopie             As Integer
Public m_bModifFournisseurBC     As Boolean
Private m_sLiaison               As String
Private m_bExtra                 As Boolean
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

40      woups "frmProjSoumElec", "PeutFermer", Err, Erl
End Function

Private Sub InitialiserVariables(ByVal sNoProjSoum As String)

5       On Error GoTo AfficherErreur

        'Initialisation des variables comprises dans la configuration
10      Dim rstConfig   As ADODB.Recordset
15      Dim rstProjSoum As ADODB.Recordset

20      Set rstProjSoum = New ADODB.Recordset

25      If m_eType = TYPE_PROJET Then
30        Call rstProjSoum.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)
35      Else
40        Call rstProjSoum.Open("SELECT * FROM GRB_SoumissionElec WHERE IDSoumission = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)
45      End If

50      If Not rstProjSoum.EOF Then
55        m_sProfit = rstProjSoum.Fields("Profit")
60        m_sCommission = rstProjSoum.Fields("Commission")
65        m_sImprevue = rstProjSoum.Fields("Imprevue")
70      Else
75        Set rstConfig = New ADODB.Recordset

80        Call rstConfig.Open("SELECT * FROM GRB_Config", g_connData, adOpenDynamic, adLockOptimistic)

85        m_sProfit = rstConfig.Fields("ProfitElec")
90        m_sCommission = rstConfig.Fields("Commission")
95        m_sImprevue = rstConfig.Fields("Imprévus")

100       Call rstConfig.Close
105       Set rstConfig = Nothing
110     End If

115     Call rstProjSoum.Close
120     Set rstProjSoum = Nothing

125     Exit Sub

AfficherErreur:

130     woups "frmProjSoumElec", "InitialiserVariables", Err, Erl
End Sub

Private Sub ActiverBoutonsGroupe()

5       On Error GoTo AfficherErreur

        'Activation des boutons d'après le groupe
10      Dim bModif As Boolean

        'Si l'utilisateur a le droit d'affichage sur les projets et les soumissions
15      If g_bAffichageProjetsElec = True And g_bAffichageSoumissionsElec = True Then
          'On affiche cmbChoix
20        cmbChoix.Visible = True

          'Cette variable sert à savoir si l'utilisateur a le droit de voir le combo
25        m_bComboChoix = True

          'Type d'affichage
30        m_eType = TYPE_PROJET

          'Champs pour la modification
35        bModif = g_bModificationProjetsElec
40      Else
          'On cache cmbChoix
45        cmbChoix.Visible = False

          'Cette variable sert à savoir si l'utilisateur a le droit de voir le combo
50        m_bComboChoix = False

          'Si l'utilisateur a le droit d'affichage sur les projets
55        If g_bAffichageProjetsElec = True Then
            'Le seul choix possible est Projet
60          txtChoix.Text = "Projet"

            'Le type d'affichage
65          m_eType = TYPE_PROJET

            'Champs pour la modification
70          bModif = g_bModificationProjetsElec
75        Else
            'Le seul choix possible est Soumission
80          txtChoix.Text = "Soumission"

            'Type d'affichage
85          m_eType = TYPE_SOUMISSION

            'Champs pour la modification
90          bModif = g_bModificationSoumissionsElec
95        End If
100     End If

105     m_bModifProj = g_bModificationProjetsElec
110     m_bModifSoum = g_bModificationSoumissionsElec
115     m_bModifBonCommande = g_bModificationBC
120     m_bSupprimer = g_bSuppressionProjets

125     Cmdajouter.Enabled = bModif
130     cmdsupprimer.Enabled = bModif
135     cmdModifier.Enabled = bModif
140     cmdCopier.Enabled = bModif
145     cmdCreerProjet.Enabled = bModif
150     cmdBonCommande.Enabled = m_bModifBonCommande
155     cmdImprimer.Enabled = bModif
160     cmdDemande.Enabled = bModif
165     cmdAnglaisFrancais.Enabled = bModif
170     cmdExtra.Enabled = bModif
175     cmdSupprimerFRS.Visible = g_bModificationCatalogueElec
180     cmdRetour.Enabled = g_bModificationRetourMarchandise
185     cmdReception.Enabled = g_bModificationReception

190     Exit Sub

AfficherErreur:

195     woups "frmProjSoumElec", "ActiverBoutonsGroupe", Err, Erl
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

60      woups "frmProjSoumElec", "AfficherProjSoum", Err, Erl
End Sub

Private Sub AfficherControles(ByVal eMode As enumMode)

5      On Error GoTo AfficherErreur

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
85      Dim bDate            As Boolean
90      Dim bTexte           As Boolean
95      Dim bCreerProjet     As Boolean
100     Dim bHistorique      As Boolean
105     Dim bCopier          As Boolean
110     Dim bBonCommande     As Boolean
115     Dim bTri             As Boolean
120     Dim bDemande         As Boolean
125     Dim bExtra           As Boolean
130     Dim bCatalogue       As Boolean
135     Dim bBrowseChemin    As Boolean
140     Dim bInutile         As Boolean
145     Dim bMauvaisPrix     As Boolean
150     Dim bRapportFact     As Boolean
155     Dim bDateFacture     As Boolean
160     Dim bSortiMagasin    As Boolean
165     Dim bRetour          As Boolean
170     Dim bForfait         As Boolean
175     Dim bExporter        As Boolean
180     Dim bReception       As Boolean
185     Dim bAnglaisFrancais As Boolean
190     Dim bRechercheClient As Boolean
  
195     m_eMode = eMode
  
200     Select Case eMode
          Case MODE_AJOUT_MODIF:
205         bEnregistrer = True
210         bAnnuler = True

215         bSection = True
220         bPieces = True
225         bTexte = True
230         bTri = True

235         If (m_eType = TYPE_SOUMISSION) Or (m_eType = TYPE_PROJET And Mid$(txtNoProjSoum.Text, 3, 1) <> "3") Then
240           bCmbClient = True
245           bCmbContact = True
250           bRechercheClient = True
255         End If

260         bCmbTransport = True
265         bDate = True
270         bCatalogue = True
275         bBrowseChemin = True
280         bMauvaisPrix = True
285         bForfait = True

290         If m_eType = TYPE_PROJET Then
295           bInutile = True

300           If g_bModificationReception = True Then
305             bSortiMagasin = True
310           End If

315           If g_bModificationFacturation = True Then
320             bDateFacture = True
325           End If
330         End If
      
335       Case MODE_INACTIF:
340         bModifier = True
345         bFermer = True
350         bImprimer = True
355         bCmbProjSoum = True
360         bCmbChoix = True
365         bCmbOuvertFerme = True
370         bHistorique = True
375         bDemande = True
380         bExporter = True
385         bAnglaisFrancais = True
390         bAjouter = True
         
395         If m_eType = TYPE_PROJET Then
400           bBonCommande = True
405           bExtra = True

410           If g_bModificationRetourMarchandise = True Then
415             bRetour = True
420           End If

425           If g_bModificationFacturation = True Then
430             bRapportFact = True
435           End If

440           If g_bModificationReception = True Then
445             bReception = True
450           End If
             
455           If m_bSupprimer = True Then
460             bSupprimer = True
465           End If
470         Else
475           bSupprimer = True
480           bCopier = True
       
485          If VerifierSiDejaProjet = False Then
490             bCreerProjet = True
495           End If
500         End If
505     End Select
  
510     Cmdajouter.Visible = bAjouter
515     cmdModifier.Visible = bModifier
520     cmdsupprimer.Visible = bSupprimer
525     cmdEnregistrer.Visible = bEnregistrer
530     cmdAnnuler.Visible = bAnnuler
535     Cmdfermer.Visible = bFermer
540     cmdImprimer.Visible = bImprimer
545     cmdRapportFACT.Visible = bRapportFact
550     cmdDate.Visible = bDate
555     cmdTexte.Visible = bTexte
560     cmdHistorique.Visible = bHistorique
565     cmdCopier.Visible = bCopier
570     cmdBonCommande.Visible = bBonCommande
575     cmdCreerProjet.Visible = bCreerProjet
580     cmdDemande.Visible = bDemande
585     cmdExtra.Visible = bExtra
590     cmdCatalogue.Visible = bCatalogue
595     cmdBrowse.Visible = bBrowseChemin
600     cmdMaterielInutile.Visible = bInutile
605     cmdMauvaisPrix.Visible = bMauvaisPrix
610     cmdSortieMagasin.Visible = bSortiMagasin
615     cmdRetour.Visible = bRetour
620     cmdForfait.Visible = bForfait
625     cmdEffacerForfait.Visible = bForfait
630     cmdExport.Visible = bExporter
635     cmdReception.Visible = bReception
640     cmdAnglaisFrancais.Visible = bAnglaisFrancais

645     lblDateFacturation.Visible = bDateFacture
650     txtDateFacturation.Visible = bDateFacture
655     cmdDateFacturation.Visible = bDateFacture
   
660     cmbclient.Visible = bCmbClient
665     txtClient.Visible = Not bCmbClient

670     cmbContact.Visible = bCmbContact
675     txtcontact.Visible = Not bCmbContact
  
680     cmbtransport.Visible = bCmbTransport
685     txtTransport.Visible = Not bCmbTransport
  
        'Si on a le droit d'afficher le combo
690     If m_bComboChoix = True Then
695       cmbChoix.Visible = bCmbChoix
700       txtChoix.Visible = Not bCmbChoix
705     End If

710     cmbOuvertFerme.Visible = bCmbOuvertFerme
  
715     cmbProjSoum.Visible = bCmbProjSoum
720     txtNoProjSoum.Visible = Not bCmbProjSoum
  
725     lblSections.Visible = bSection
730     cmbSections.Visible = bSection
735     cmdAjouterSection.Visible = bSection

740     lblPiece.Visible = bPieces
745     cmbPieces.Visible = bPieces
750     lvwPieces.Visible = bPieces

755     lblTri.Visible = bTri
760     cmbTri.Visible = bTri
765     cmdTri.Visible = bTri
770     cmdRafraichir.Visible = bTri
    
775     fraPrix.Visible = m_bDroitPrix

780     cmdRechercherClient.Visible = bRechercheClient

785     Exit Sub

AfficherErreur:

790     woups "frmProjSoumElec", "AfficherControles", Err, Erl
End Sub

Private Sub cmbChoix_Click()

5       On Error GoTo AfficherErreur

10      Dim bModif          As Boolean
15      Dim iCmbOuvertFerme As Integer
    
20      Screen.MousePointer = vbHourglass
      
25      txtChoix.Text = cmbChoix.Text

        'Mets les CheckBoxes sur le ListView
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
240     cmdAnglaisFrancais.Enabled = bModif
245     cmdDemande.Enabled = bModif
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

300     woups "frmProjSoumElec", "cmbChoix_Click", Err, Erl
End Sub

Private Sub cmbclient_Click()

5       On Error GoTo AfficherErreur

        'Rempli le combo des contacts selon le client choisi
10      Call RemplirComboContacts

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElec", "cmbclient_Click", Err, Erl
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

20      woups "frmProjSoumElec", "cmbPieces_Click", Err, Erl
End Sub

Private Sub cmbProjSoum_Click()

5       On Error GoTo AfficherErreur

10      Dim rstProjSoum As ADODB.Recordset
15      Dim rstOuvert   As ADODB.Recordset
20      Dim iCompteur   As Integer
25      Dim sNomClient  As String
30      Dim sNomContact As String
35      Dim sNumero     As String
40      Dim sTransport  As String
45      Dim bTrouve     As Boolean
  
50      Screen.MousePointer = vbHourglass

55      m_bRecherchePiece = False
60      m_bChangementFRS = False
65      m_bPieceInutile = False

70      If cmbProjSoum.Text <> "" Then
75        sNumero = txtNoProjSoum.Text

80        txtNoProjSoum.Text = cmbProjSoum.Text

85        Call InitialiserVariables(txtNoProjSoum.Text)

90        If m_bEnregistrement = False Then
95          m_eLangage = FRANCAIS

100         cmdAnglaisFrancais.Caption = "Anglais"
105       End If
  
110       Set rstProjSoum = New ADODB.Recordset
  
115       If m_eType = TYPE_SOUMISSION Then
120         Call rstProjSoum.Open("SELECT * FROM GRB_SoumissionElec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
125       Else
130         Call rstProjSoum.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
135       End If

140       If rstProjSoum.Fields("Modification") = True And rstProjSoum.Fields("Par") = g_sEmploye Then
145         cmdReset.Visible = True
150       End If

155       Call InitialiserTempsTaux(False)

160       If m_eType = TYPE_SOUMISSION Then
            'Si la soumission n'est pas assigné à un projet
165         If VerifierSiDejaProjet = False Then
              'On affiche le bouton cmdCreerProjet
170           cmdCreerProjet.Visible = True
175         Else
180           cmdCreerProjet.Visible = False
185         End If
190       End If
  
          'Rempli les valeurs de la soumission ou du projet sélectionné
195       Call RemplirProjSoum

          'Le temps calculé dans le projet est le temps réel, c'est pourquoi il faut le recalculer
          'puisque le temps varie souvent
200       If m_eType = TYPE_PROJET Then
205         Set rstOuvert = New ADODB.Recordset

210         Call rstOuvert.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

215         If rstOuvert.Fields("Ouvert") = True Then
220           m_bModeAffichage = False

225           Call CalculerPrix

230           m_bModeAffichage = True

235           rstProjSoum.Fields("total_Commission") = txtCommission.Text
240           rstProjSoum.Fields("Total_Profit") = txtProfit.Text
245           rstProjSoum.Fields("PrixTotal") = txtPrixTotal.Text
250           rstProjSoum.Fields("Total_piece") = txtTotalPieces.Text
255           rstProjSoum.Fields("total_imprevue") = txtImprevus.Text
260           rstProjSoum.Fields("Total_Temps") = txtTotalTemps.Text

265           Call rstProjSoum.Update
270         End If
275       End If
  
280       Call rstProjSoum.Close

285       sNomClient = txtClient.Text
290       sNomContact = txtcontact.Text
295       sTransport = txtTransport.Text
    
          'Pour choisir le bon client dans le combo des clients
300       For iCompteur = 0 To cmbclient.ListCount - 1
305         If cmbclient.LIST(iCompteur) = sNomClient Then
310           cmbclient.ListIndex = iCompteur

315           bTrouve = True
    
320           Exit For
325         End If
330       Next
    
335       If bTrouve = False Then
340         Call RemplirComboClients(vbNullString)

345         For iCompteur = 0 To cmbclient.ListCount - 1
350           If cmbclient.LIST(iCompteur) = sNomClient Then
355             cmbclient.ListIndex = iCompteur

360             Exit For
365           End If
370         Next
375       End If
    
          'Pour choisir le bon contact dans le combo des contacts
380       For iCompteur = 0 To cmbContact.ListCount - 1
385         If cmbContact.LIST(iCompteur) = sNomContact Then
390           cmbContact.ListIndex = iCompteur
      
395           Exit For
400         End If
405       Next
  
          'Pour choisir le bon transport dans le combo des transports
410       For iCompteur = 0 To cmbtransport.ListCount - 1
415         If cmbtransport.LIST(iCompteur) = sTransport Then
420           cmbtransport.ListIndex = iCompteur
      
425           Exit For
430         End If
435       Next
440     End If

445     Call CalculerPrixReception

450     If m_eType = TYPE_PROJET Then
455       rstProjSoum.CursorLocation = adUseServer

460       Call rstProjSoum.Open("SELECT PrixRéception FROM GRB_ProjetElec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
    
465       rstProjSoum.Fields("PrixRéception") = txtPrixReception.Text

470       Call rstProjSoum.Update

475       Call rstProjSoum.Close
480     End If

485     If m_bSansTemps = True Then
490       tmrTemps.Enabled = True
495     Else
500       tmrTemps.Enabled = False
505       lblPasTemps.Visible = False
510     End If
  
515     Set rstProjSoum = Nothing
  
520     Screen.MousePointer = vbDefault

525     Exit Sub

AfficherErreur:

530     woups "frmProjSoumElec", "cmbProjSoum_Click", Err, Erl, txtNoProjSoum.Text)
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
40        m_sTempsFabrication = "0"
45        m_sTempsAssemblage = "0"
50        m_sTempsProgInterface = "0"
55        m_sTempsProgAutomate = "0"
60        m_sTempsProgRobot = "0"
65        m_sTempsVision = "0"
70        m_sTempsTest = "0"
75        m_sTempsInstallation = "0"
80        m_sTempsMiseService = "0"
85        m_sTempsFormation = "0"
90        m_sTempsGestion = "0"
95        m_sTempsShipping = "0"

100       m_sNbrePersonne = "0"
105       m_sTempsHebergement = "0"
110       m_sTempsRepas = "0"
115       m_sTempsTransport = "0"
120       m_sTempsUniteMobile = "0"
125       m_sPrixEmballage = "0"
130       m_sTauxHebergement1 = "0"
135       m_sTauxHebergement2 = "0"
140       m_sTauxRepas = "0"
145       m_sTauxTransport = "0"
150       m_sTauxUniteMobile = "0"

155       m_sTauxDessin = "0"
160       m_sTauxFabrication = "0"
165       m_sTauxAssemblage = "0"
170       m_sTauxProgInterface = "0"
175       m_sTauxProgAutomate = "0"
180       m_sTauxProgRobot = "0"
185       m_sTauxVision = "0"
190       m_sTauxTest = "0"
195       m_sTauxInstallation = "0"
200       m_sTauxMiseService = "0"
205       m_sTauxFormation = "0"
210       m_sTauxGestion = "0"
215       m_sTauxShipping = "0"
220     Else
225       If m_eType = TYPE_PROJET Then
230         sTable = "GRB_ProjetElec"
235         sChamps = "IDProjet"
240       Else
245         sTable = "GRB_SoumissionElec"
250         sChamps = "IDSoumission"
255       End If

260       Set rstProjSoum = New ADODB.Recordset

265       Call rstProjSoum.Open("SELECT * FROM " & sTable & " WHERE " & sChamps & " = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

270       If m_eType = TYPE_SOUMISSION Then
275         If Not IsNull(rstProjSoum.Fields("TempsDessin")) Then
280           m_sTempsDessin = rstProjSoum.Fields("TempsDessin")
285         Else
290           m_sTempsDessin = "0"
295         End If

300         If Not IsNull(rstProjSoum.Fields("TempsFabrication")) Then
305           m_sTempsFabrication = rstProjSoum.Fields("TempsFabrication")
310         Else
315           m_sTempsFabrication = "0"
320         End If

325         If Not IsNull(rstProjSoum.Fields("TempsAssemblage")) Then
330           m_sTempsAssemblage = rstProjSoum.Fields("TempsAssemblage")
335         Else
340           m_sTempsAssemblage = "0"
345         End If

350         If Not IsNull(rstProjSoum.Fields("TempsProgInterface")) Then
355           m_sTempsProgInterface = rstProjSoum.Fields("TempsProgInterface")
360         Else
365           m_sTempsProgInterface = "0"
370         End If

375         If Not IsNull(rstProjSoum.Fields("TempsProgAutomate")) Then
380           m_sTempsProgAutomate = rstProjSoum.Fields("TempsProgAutomate")
385         Else
390           m_sTempsProgAutomate = "0"
395         End If

400         If Not IsNull(rstProjSoum.Fields("TempsProgRobot")) Then
405           m_sTempsProgRobot = rstProjSoum.Fields("TempsProgRobot")
410         Else
415           m_sTempsProgRobot = "0"
420         End If

425         If Not IsNull(rstProjSoum.Fields("TempsVision")) Then
430           m_sTempsVision = rstProjSoum.Fields("TempsVision")
435         Else
440           m_sTempsVision = "0"
445         End If

450         If Not IsNull(rstProjSoum.Fields("TempsTest")) Then
455           m_sTempsTest = rstProjSoum.Fields("TempsTest")
460         Else
465           m_sTempsTest = "0"
470         End If

475         If Not IsNull(rstProjSoum.Fields("TempsInstallation")) Then
480           m_sTempsInstallation = rstProjSoum.Fields("TempsInstallation")
485         Else
490           m_sTempsInstallation = "0"
495         End If

500         If Not IsNull(rstProjSoum.Fields("TempsMiseService")) Then
505           m_sTempsMiseService = rstProjSoum.Fields("TempsMiseService")
510         Else
515           m_sTempsMiseService = "0"
520         End If

525         If Not IsNull(rstProjSoum.Fields("TempsFormation")) Then
530           m_sTempsFormation = rstProjSoum.Fields("TempsFormation")
535         Else
540           m_sTempsFormation = "0"
545         End If

550         If Not IsNull(rstProjSoum.Fields("TempsGestion")) Then
555           m_sTempsGestion = rstProjSoum.Fields("TempsGestion")
560         Else
565           m_sTempsGestion = "0"
570         End If

575         If Not IsNull(rstProjSoum.Fields("TempsShipping")) Then
580           m_sTempsShipping = rstProjSoum.Fields("TempsShipping")
585         Else
590           m_sTempsShipping = "0"
595         End If
600       Else
605         Call InitialiserTempsReel
610       End If

615       If Not IsNull(rstProjSoum.Fields("TauxDessin")) Then
620         m_sTauxDessin = rstProjSoum.Fields("TauxDessin")
625       Else
630         m_sTauxDessin = "0"
635       End If

640       If Not IsNull(rstProjSoum.Fields("TauxFabrication")) Then
645         m_sTauxFabrication = rstProjSoum.Fields("TauxFabrication")
650       Else
655         m_sTauxFabrication = "0"
660       End If

665       If Not IsNull(rstProjSoum.Fields("TauxAssemblage")) Then
670         m_sTauxAssemblage = rstProjSoum.Fields("TauxAssemblage")
675       Else
680         m_sTauxAssemblage = "0"
685       End If

690       If Not IsNull(rstProjSoum.Fields("TauxProgInterface")) Then
695         m_sTauxProgInterface = rstProjSoum.Fields("TauxProgInterface")
700       Else
705         m_sTauxProgInterface = "0"
710       End If

715       If Not IsNull(rstProjSoum.Fields("TauxProgAutomate")) Then
720         m_sTauxProgAutomate = rstProjSoum.Fields("TauxProgAutomate")
725       Else
730         m_sTauxProgAutomate = "0"
735       End If

740       If Not IsNull(rstProjSoum.Fields("TauxProgRobot")) Then
745         m_sTauxProgRobot = rstProjSoum.Fields("TauxProgRobot")
750       Else
755         m_sTauxProgRobot = "0"
760       End If

765       If Not IsNull(rstProjSoum.Fields("TauxVision")) Then
770         m_sTauxVision = rstProjSoum.Fields("TauxVision")
775       Else
780         m_sTauxVision = "0"
785       End If

790       If Not IsNull(rstProjSoum.Fields("TauxTest")) Then
795         m_sTauxTest = rstProjSoum.Fields("TauxTest")
800       Else
805         m_sTauxTest = "0"
810       End If

815       If Not IsNull(rstProjSoum.Fields("TauxInstallation")) Then
820         m_sTauxInstallation = rstProjSoum.Fields("TauxInstallation")
825       Else
830         m_sTauxInstallation = "0"
835       End If

840       If Not IsNull(rstProjSoum.Fields("TauxMiseService")) Then
845         m_sTauxMiseService = rstProjSoum.Fields("TauxMiseService")
850       Else
855         m_sTauxMiseService = "0"
860       End If

865       If Not IsNull(rstProjSoum.Fields("TauxFormation")) Then
870         m_sTauxFormation = rstProjSoum.Fields("TauxFormation")
875       Else
880         m_sTauxFormation = "0"
885       End If

890       If Not IsNull(rstProjSoum.Fields("TauxGestion")) Then
895         m_sTauxGestion = rstProjSoum.Fields("TauxGestion")
900       Else
905         m_sTauxGestion = "0"
910       End If

915       If Not IsNull(rstProjSoum.Fields("TauxShipping")) Then
920         m_sTauxShipping = rstProjSoum.Fields("TauxShipping")
925       Else
930         m_sTauxShipping = "0"
935       End If

940       If m_eType = TYPE_PROJET Then
945         m_sNbrePersonne = "0"
950         m_sTempsHebergement = "0"
955         m_sTempsRepas = "0"
960         m_sTempsTransport = "0"
965         m_sTempsUniteMobile = "0"
970       Else
975         If Not IsNull(rstProjSoum.Fields("NbrePersonne")) Then
980           m_sNbrePersonne = rstProjSoum.Fields("NbrePersonne")
985         Else
990           m_sNbrePersonne = "0"
995         End If
    
1000        If Not IsNull(rstProjSoum.Fields("TempsHebergement")) Then
1005          m_sTempsHebergement = rstProjSoum.Fields("TempsHebergement")
1010        Else
1015          m_sTempsHebergement = "0"
1020        End If

1025        If Not IsNull(rstProjSoum.Fields("TempsRepas")) Then
1030          m_sTempsRepas = rstProjSoum.Fields("TempsRepas")
1035        Else
1040          m_sTempsRepas = "0"
1045        End If

1050        If Not IsNull(rstProjSoum.Fields("TempsTransport")) Then
1055          m_sTempsTransport = rstProjSoum.Fields("TempsTransport")
1060        Else
1065          m_sTempsTransport = "0"
1070        End If

1075        If Not IsNull(rstProjSoum.Fields("TempsUniteMobile")) Then
1080          m_sTempsUniteMobile = rstProjSoum.Fields("TempsUniteMobile")
1085        Else
1090          m_sTempsUniteMobile = "0"
1095        End If
1100      End If

1105      If Not IsNull(rstProjSoum.Fields("PrixEmballage")) Then
1110        m_sPrixEmballage = rstProjSoum.Fields("PrixEmballage")
1115      Else
1120        m_sPrixEmballage = "0"
1125      End If

1130      If m_eType = TYPE_PROJET Then
1135        m_sTauxHebergement1 = "0"
1140        m_sTauxHebergement2 = "0"
1145        m_sTauxRepas = "0"
1150        m_sTauxTransport = "0"
1155        m_sTauxUniteMobile = "0"
1160      Else
1165        If Not IsNull(rstProjSoum.Fields("TauxHebergement1")) Then
1170          m_sTauxHebergement1 = rstProjSoum.Fields("TauxHebergement1")
1175        Else
1180          m_sTauxHebergement1 = "0"
1185        End If

1190        If Not IsNull(rstProjSoum.Fields("TauxHebergement2")) Then
1195          m_sTauxHebergement2 = rstProjSoum.Fields("TauxHebergement2")
1200        Else
1205          m_sTauxHebergement2 = "0"
1210        End If

1215        If Not IsNull(rstProjSoum.Fields("TauxRepas")) Then
1220          m_sTauxRepas = rstProjSoum.Fields("TauxRepas")
1225        Else
1230          m_sTauxRepas = "0"
1235        End If

1240        If Not IsNull(rstProjSoum.Fields("TauxTransport")) Then
1245          m_sTauxTransport = rstProjSoum.Fields("TauxTransport")
1250        Else
1255          m_sTauxTransport = "0"
1260        End If

1265        If Not IsNull(rstProjSoum.Fields("TauxUniteMobile")) Then
1270          m_sTauxUniteMobile = rstProjSoum.Fields("TauxUniteMobile")
1275        Else
1280          m_sTauxUniteMobile = "0"
1285        End If
1290      End If

1295      Call rstProjSoum.Close
1300      Set rstProjSoum = Nothing
1305    End If

1310    Exit Sub

AfficherErreur:

1315    woups "frmProjSoumElec", "InitialiserTempsTaux", Err, Erl
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
90      m_sTempsFabrication = "0"
95      m_sTempsAssemblage = "0"
100     m_sTempsProgInterface = "0"
105     m_sTempsProgAutomate = "0"
110     m_sTempsProgRobot = "0"
115     m_sTempsVision = "0"
120     m_sTempsTest = "0"
125     m_sTempsInstallation = "0"
130     m_sTempsMiseService = "0"
135     m_sTempsFormation = "0"
140     m_sTempsGestion = "0"
145     m_sTempsShipping = "0"

150     Do While Not rstPunch.EOF
155       If Not IsNull(rstPunch.Fields("Total")) Then
160         Select Case rstPunch.Fields("Type")
              Case "Dessin":        m_sTempsDessin = Round(rstPunch.Fields("Total"), 2)
165           Case "Fabrication":   m_sTempsFabrication = Round(rstPunch.Fields("Total"), 2)
170           Case "Assemblage":    m_sTempsAssemblage = Round(rstPunch.Fields("Total"), 2)
175           Case "ProgInterface": m_sTempsProgInterface = Round(rstPunch.Fields("Total"), 2)
180           Case "ProgAutomate":  m_sTempsProgAutomate = Round(rstPunch.Fields("Total"), 2)
185           Case "ProgRobot":     m_sTempsProgRobot = Round(rstPunch.Fields("Total"), 2)
190           Case "Vision":        m_sTempsVision = Round(rstPunch.Fields("Total"), 2)
195           Case "Test":          m_sTempsTest = Round(rstPunch.Fields("Total"), 2)
200           Case "Installation":  m_sTempsInstallation = Round(rstPunch.Fields("Total"), 2)
205           Case "MiseService":   m_sTempsMiseService = Round(rstPunch.Fields("Total"), 2)
210           Case "Formation":     m_sTempsFormation = Round(rstPunch.Fields("Total"), 2)
215           Case "Gestion":       m_sTempsGestion = Round(rstPunch.Fields("Total"), 2)
220           Case "Shipping":      m_sTempsShipping = Round(rstPunch.Fields("Total"), 2)
225         End Select
230       End If

235       Call rstPunch.MoveNext
240     Loop

245     Call rstPunch.Close

250     Set rstPunch = Nothing

255     Exit Sub

AfficherErreur:

260     woups "frmProjSoumElecTemps", "AfficherTempsReels", Err, Erl
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

50      woups "frmProjSoumElec", "cmbProjSoum_KeyUp", Err, Erl
End Sub

Private Sub cmdAjouterSection_Click()

5       On Error GoTo AfficherErreur

        'Affiche le form frmSoumissionSection
10      Call OuvrirForm(frmSoumissionSectionElec, True)

        'Après que l'utilisateur a refermé le form, on rafraichi le
        'contenu du combo
15      Call RemplirComboSections

20      Call UpdateOrdre

25      Exit Sub

AfficherErreur:

30      woups "frmProjSoumElec", "cmdAjouterSection_Click", Err, Erl
End Sub

Private Sub cmdAnglaisFrancais_Click()

5       On Error GoTo AfficherErreur

10      Screen.MousePointer = vbHourglass

15      If cmdAnglaisFrancais.Caption = "Anglais" Then
20        m_eLangage = ANGLAIS
    
25        cmdAnglaisFrancais.Caption = "Français"
30      Else
35        m_eLangage = FRANCAIS
    
40        cmdAnglaisFrancais.Caption = "Anglais"
45      End If

50      Call UpdateDescription
  
55      Call RemplirComboSections
    
60      Call UpdateOrdre
  
65      Screen.MousePointer = vbDefault

70      Exit Sub

AfficherErreur:

75      woups "frmProjSoumElec", "cmdAnglaisFrancais_Click", Err, Erl
End Sub

Private Sub UpdateDescription()

5       On Error GoTo AfficherErreur

10      Dim rstProjSoum  As ADODB.Recordset
15      Dim rstPieceElec As ADODB.Recordset

20      Set rstProjSoum = New ADODB.Recordset
25      Set rstPieceElec = New ADODB.Recordset

30      If m_eType = TYPE_PROJET Then
35        Call rstProjSoum.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
40      Else
45        Call rstProjSoum.Open("SELECT * FROM GRB_Soumission_Pieces WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
50      End If

55      Do While Not rstProjSoum.EOF
60        Call rstPieceElec.Open("SELECT * FROM GRB_CatalogueElec WHERE PIECE = '" & rstProjSoum.Fields("NumItem") & "'", g_connData, adOpenDynamic, adLockOptimistic)

65        rstProjSoum.Fields("Desc_Fr") = rstPieceElec.Fields("DESC_FR")
70        rstProjSoum.Fields("Desc_En") = rstPieceElec.Fields("DESC_EN")

75        Call rstProjSoum.Update

80        Call rstPieceElec.Close

85        Call rstProjSoum.MoveNext
90      Loop

95      Set rstPieceElec = Nothing

100     Call rstProjSoum.Close
105     Set rstProjSoum = Nothing

110     Call RemplirListViewProjSoum(txtNoProjSoum.Text)

115     Exit Sub

AfficherErreur:

120     woups "frmProjSoumElec", "UpdateDescription", Err, Erl
End Sub

Private Sub cmdAnnuler_Click()

5       On Error GoTo AfficherErreur
  
10      fraPieceTrouve.Visible = False
15      frafournisseur.Visible = False
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

75      woups "frmProjSoumElec", "cmdAnnuler_Click", Err, Erl
End Sub

Private Sub cmdAnnulerCommentaire_Click()

5       On Error GoTo AfficherErreur

10      fraCommentaire.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElec", "cmdAnnulerCommentaire_Click", Err, Erl
End Sub

Private Sub cmdAnnulerDateRequise_Click()

5       On Error GoTo AfficherErreur

10      fraDateRequise.Visible = False

15      m_bMonthViewHasFocus = False

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumElec", "cmdAnnulerDateRequise_Click", Err, Erl
End Sub

Private Sub cmdAnnulerDateRequise_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
5       On Error GoTo AfficherErreur

10      If m_bMonthViewHasFocus = True Then
15        Call cmdAnnulerDateRequise_Click
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmProjSoumElec", "cmdAnnulerDateRequise_MouseUp", Err, Erl
End Sub

Private Sub cmdExport_Click()

5       On Error GoTo AfficherErreur

10      Dim rstProjSoum As ADODB.Recordset

15      Set rstProjSoum = New ADODB.Recordset

20      Call g_connData.Execute("DELETE * FROM GRB_impression_listepiece")

25      If m_eType = TYPE_PROJET Then
30        Call rstProjSoum.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
35      Else
40        Call rstProjSoum.Open("SELECT * FROM GRB_SoumissionElec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
45      End If

50      Call ExporterListePieces(rstProjSoum)

55      Call rstProjSoum.Close
60      Set rstProjSoum = Nothing

65      Exit Sub

AfficherErreur:

70      woups "frmProjSoumElec", "cmdExport_Click", Err, Erl
End Sub

Private Sub ExporterListePieces(ByVal rstProjSoum As ADODB.Recordset)

5       On Error GoTo AfficherErreur

        'Impression de la feuille de la liste des pièces
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

        'Ouverture du recordset
110     If m_eType = TYPE_PROJET Then
115       sNoProjet = rstProjSoum.Fields("IDProjet")
120       sNoSoumission = rstProjSoum.Fields("IDSoumission")

125       Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoProjet & "' AND Type = 'E' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
130     Else
135       sNoProjet = vbNullString
140       sNoSoumission = rstProjSoum.Fields("IDSoumission")

145       Call rstPiece.Open("SELECT * FROM GRB_Soumission_Pieces WHERE IDSoumission = '" & sNoSoumission & "' AND Type = 'E' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
150     End If

155     Do While Not rstPiece.EOF
160       If rstPiece.Fields("Visible") = True Then
165         bAjouterSection = True
170         bAjouterSousSection = True
175         bAjouterPiece = True

180         rstImpListePiece.CursorLocation = adUseClient

185         Call rstImpListePiece.Open("SELECT * FROM GRB_Impression_ListePiece WHERE IDSection = '" & rstPiece.Fields("IDSection") & "'", g_connData, adOpenDynamic, adLockOptimistic)

190         If Not rstImpListePiece.EOF Then
195           bAjouterSection = False

200           Do While Not rstImpListePiece.EOF
205             If rstImpListePiece.Fields("NomSousSection") = rstPiece.Fields("SousSection") Then
210               bAjouterSousSection = False

215               If rstPiece.Fields("NumItem") <> "Texte" And rstPiece.Fields("NumItem") <> "Text" Then
220                 If rstImpListePiece.Fields("NumItem") = rstPiece.Fields("NumItem") Then
225                   bAjouterPiece = False

230                   rstImpListePiece.Fields("Qté") = Replace(CDbl(rstImpListePiece.Fields("Qté")) + CDbl(rstPiece.Fields("Qté")), ".", ",")

235                   If Not IsNull(rstImpListePiece.Fields("ID")) Then
240                     If rstImpListePiece.Fields("ID") <> "" Then
245                       rstImpListePiece.Fields("ID") = rstImpListePiece.Fields("ID") & ", " & rstPiece.Fields("ID")
250                     Else
255                       rstImpListePiece.Fields("ID") = rstPiece.Fields("ID")
260                     End If
265                   Else
270                     rstImpListePiece.Fields("ID") = rstPiece.Fields("ID")
275                   End If

280                   Call rstImpListePiece.Update

285                   If rstImpListePiece.Fields("Qté") = 0 Then
290                     Call rstImpListePiece.Delete

295                     rstImpListePiece.Filter = "NomSousSection = '" & Replace(rstPiece.Fields("SousSection"), "'", "''") & "'"

300                     If rstImpListePiece.RecordCount = 1 Then
305                       Call rstImpListePiece.Delete

310                       rstImpListePiece.Filter = "IDSection = '" & rstPiece.Fields("IDSection") & "'"

315                       If rstImpListePiece.RecordCount = 1 Then
320                         Call rstImpListePiece.Delete
325                       End If
330                     End If

335                     rstImpListePiece.Filter = ""
340                   End If

345                   Exit Do
350                 End If
355               Else
360                 Exit Do
365               End If
370             End If

375             Call rstImpListePiece.MoveNext
380           Loop
385         End If

390         If bAjouterSection = True Then
395           If m_eLangage = ANGLAIS Then
400             sSection = "NomSectionEN"
405           Else
410             sSection = "NomSectionFR"
415           End If

420           Call rstTemp.Open("SELECT " & sSection & " FROM GRB_SoumProjSectionElec WHERE IDSection = " & rstPiece.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)

              'Ajoute la section dans la liste de pièces
425           Call rstImpListePiece.AddNew

430           rstImpListePiece.Fields("NoLigne") = iCompteurPiece
435           rstImpListePiece.Fields("IDSoumission") = sNoSoumission

440           If Not IsNull(rstTemp.Fields(sSection)) Then
445             rstImpListePiece.Fields("Section") = rstTemp.Fields(sSection)
450           Else
455             rstImpListePiece.Fields("Section") = " "
460           End If

465           rstImpListePiece.Fields("IDSection") = rstPiece.Fields("IDSection")

470           Call rstImpListePiece.Update

475           iCompteurPiece = iCompteurPiece + 1

480           Call rstTemp.Close
485         End If

490         If bAjouterSousSection = True Then
495           sSousSection = rstPiece.Fields("SousSection")

500           If sSousSection = S_PAS_SOUS_SECTION Then
505             sSousSection = " "
510           End If

515           Call rstImpListePiece.AddNew

520           rstImpListePiece.Fields("NoLigne") = iCompteurPiece
525           rstImpListePiece.Fields("IDSoumission") = sNoSoumission
530           rstImpListePiece.Fields("SousSection") = sSousSection
535           rstImpListePiece.Fields("NomSousSection") = rstPiece.Fields("SousSection")
540           rstImpListePiece.Fields("IDSection") = rstPiece.Fields("IDSection")

545           Call rstImpListePiece.Update

550           iCompteurPiece = iCompteurPiece + 1
555         End If

560         If bAjouterPiece = True Then
              'Ajoute la pièce à la liste de pièces
565           Call rstImpListePiece.AddNew

570           rstImpListePiece.Fields("NoLigne") = iCompteurPiece
575           rstImpListePiece.Fields("IDSoumission") = sNoSoumission
580           rstImpListePiece.Fields("NumItem") = rstPiece.Fields("NumItem")
585           rstImpListePiece.Fields("Qté") = rstPiece.Fields("Qté")

590           If m_eLangage = ANGLAIS Then
595             rstImpListePiece.Fields("Description") = rstPiece.Fields("Desc_EN")
600           Else
605             rstImpListePiece.Fields("Description") = rstPiece.Fields("Desc_FR")
610           End If

615           rstImpListePiece.Fields("Manufact") = rstPiece.Fields("Manufact")

620           If m_eType = TYPE_PROJET Then
625             rstImpListePiece.Fields("ID") = rstPiece.Fields("ID")
630           End If

635           rstImpListePiece.Fields("IDSection") = rstPiece.Fields("IDSection")
640           rstImpListePiece.Fields("NomSousSection") = rstPiece.Fields("SousSection")

645           Call rstImpListePiece.Update

650           iCompteurPiece = iCompteurPiece + 1
655         End If

660         Call rstImpListePiece.Close
665       End If

          'Prochaine enregistrement
670       Call rstPiece.MoveNext
675     Loop

        ''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Rapport liste pièce, met dans l'ordre de ligne '
        ''''''''''''''''''''''''''''''''''''''''''''''''''
680     rstImpListePiece.CursorLocation = adUseClient

685     Call rstImpListePiece.Open("SELECT * FROM GRB_impression_Listepiece WHERE TRIM(IDSoumission) = '" & Trim$(sNoSoumission) & "' ORDER BY noligne", g_connData, adOpenDynamic, adLockOptimistic)

690     Set xlsApp = New Excel.Application

695     Set xlsWorkBook = xlsApp.Workbooks.Add

700     xlsApp.Range("A1") = "Liste de matériel ( " & txtNoProjSoum.Text & " )"
705     xlsApp.Range("A1").Font.Bold = True
710     xlsApp.Range("A1").Font.Underline = xlUnderlineStyleSingle
715     xlsApp.Range("A1").HorizontalAlignment = xlCenter
720     xlsApp.Range("A1").Font.SIZE = 14

725     Call xlsApp.Range("A1:E1").Merge

730     xlsApp.Range("A4") = "Qté"
735     xlsApp.Range("A4").Font.Bold = True
740     xlsApp.Range("A4").HorizontalAlignment = xlCenter

745     xlsApp.Range("B4") = "No. Item"
750     xlsApp.Range("B4").Font.Bold = True
755     xlsApp.Range("B4").HorizontalAlignment = xlCenter

760     xlsApp.Range("C4") = "Description"
765     xlsApp.Range("C4").Font.Bold = True
770     xlsApp.Range("C4").HorizontalAlignment = xlCenter

775     xlsApp.Range("D4") = "Manufacturier"
780     xlsApp.Range("D4").Font.Bold = True
785     xlsApp.Range("D4").HorizontalAlignment = xlCenter

790     xlsApp.Range("E4") = "#ID"
795     xlsApp.Range("E4").Font.Bold = True
800     xlsApp.Range("E4").HorizontalAlignment = xlCenter

805     xlsApp.Range("A4:E4").Borders(xlEdgeBottom).LineStyle = xlContinuous
810     xlsApp.Range("A4:E4").Borders(xlEdgeBottom).Weight = xlMedium
815     xlsApp.Range("A4:E4").Borders(xlEdgeBottom).ColorIndex = xlAutomatic

820     xlsApp.Range("A4:E4").Borders(xlInsideVertical).LineStyle = xlContinuous
825     xlsApp.Range("A4:E4").Borders(xlInsideVertical).Weight = xlMedium
830     xlsApp.Range("A4:E4").Borders(xlInsideVertical).ColorIndex = xlAutomatic

835     iCompteur = 5

840     Do While Not rstImpListePiece.EOF
845       xlsApp.Range("A" & iCompteur) = rstImpListePiece.Fields("Qté")

850       If IsNull(rstImpListePiece.Fields("Section")) Then
855         xlsApp.Range("B" & iCompteur) = rstImpListePiece.Fields("NumItem")
860       Else
865         xlsApp.Range("B" & iCompteur) = rstImpListePiece.Fields("Section")
870         xlsApp.Range("B" & iCompteur).Font.Bold = True
875       End If

880       If IsNull(rstImpListePiece.Fields("SousSection")) Then
885         xlsApp.Range("C" & iCompteur) = rstImpListePiece.Fields("Description")
890       Else
895         xlsApp.Range("C" & iCompteur) = rstImpListePiece.Fields("SousSection")
900         xlsApp.Range("C" & iCompteur).Font.Bold = True
905       End If

910       xlsApp.Range("D" & iCompteur) = rstImpListePiece.Fields("Manufact")
915       xlsApp.Range("E" & iCompteur) = rstImpListePiece.Fields("ID")

920       xlsApp.Range("A" & iCompteur & ":E" & iCompteur).Borders(xlEdgeBottom).LineStyle = xlContinuous
925       xlsApp.Range("A" & iCompteur & ":E" & iCompteur).Borders(xlEdgeBottom).Weight = xlThin
930       xlsApp.Range("A" & iCompteur & ":E" & iCompteur).Borders(xlEdgeBottom).ColorIndex = xlAutomatic

935       xlsApp.Range("A" & iCompteur & ":E" & iCompteur).Borders(xlInsideVertical).LineStyle = xlContinuous
940       xlsApp.Range("A" & iCompteur & ":E" & iCompteur).Borders(xlInsideVertical).Weight = xlThin
945       xlsApp.Range("A" & iCompteur & ":E" & iCompteur).Borders(xlInsideVertical).ColorIndex = xlAutomatic

950       Call rstImpListePiece.MoveNext

955       iCompteur = iCompteur + 1
960     Loop

965     iCompteur = iCompteur - 1

970     xlsApp.Range("A4:E" & iCompteur).Borders(xlEdgeBottom).LineStyle = xlContinuous
975     xlsApp.Range("A4:E" & iCompteur).Borders(xlEdgeBottom).Weight = xlMedium
980     xlsApp.Range("A4:E" & iCompteur).Borders(xlEdgeBottom).ColorIndex = xlAutomatic

985     xlsApp.Range("A4:E" & iCompteur).Borders(xlEdgeTop).LineStyle = xlContinuous
990     xlsApp.Range("A4:E" & iCompteur).Borders(xlEdgeTop).Weight = xlMedium
995     xlsApp.Range("A4:E" & iCompteur).Borders(xlEdgeTop).ColorIndex = xlAutomatic

1000    xlsApp.Range("A4:E" & iCompteur).Borders(xlEdgeLeft).LineStyle = xlContinuous
1005    xlsApp.Range("A4:E" & iCompteur).Borders(xlEdgeLeft).Weight = xlMedium
1010    xlsApp.Range("A4:E" & iCompteur).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

1015    xlsApp.Range("A4:E" & iCompteur).Borders(xlEdgeRight).LineStyle = xlContinuous
1020    xlsApp.Range("A4:E" & iCompteur).Borders(xlEdgeRight).Weight = xlMedium
1025    xlsApp.Range("A4:E" & iCompteur).Borders(xlEdgeRight).ColorIndex = xlAutomatic

1030    Call xlsApp.Columns("A:A").EntireColumn.AutoFit
1035    Call xlsApp.Columns("B:B").EntireColumn.AutoFit
1040    Call xlsApp.Columns("C:C").EntireColumn.AutoFit
1045    Call xlsApp.Columns("D:D").EntireColumn.AutoFit
1050    Call xlsApp.Columns("E:E").EntireColumn.AutoFit

1055    Call rstImpListePiece.Close
1060    Set rstImpListePiece = Nothing

1065    Screen.MousePointer = vbDefault

1070    sSaveAsFileName = xlsApp.GetSaveAsFilename(txtNoProjSoum.Text & ".xls", "Fichiers Excel (*.xlx), *.xls")

1075    If sSaveAsFileName <> "Faux" Then
1080      Call xlsWorkBook.SaveAs(sSaveAsFileName)
1085    End If

1090    xlsWorkBook.Saved = True

1095    Call xlsWorkBook.Close

1100    Set xlsWorkBook = Nothing

1105    Call xlsApp.Quit

1110    Set xlsApp = Nothing

1115    Set rstTemp = Nothing

1120    Exit Sub

AfficherErreur:

1125    woups "frmProjSoumElec", "ExporterListePieces", Err, Erl
End Sub

Private Sub cmdOKCommentaire_Click()

5       On Error GoTo AfficherErreur

10      lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_COMMENTAIRE) = txtcommentaire.Text

15      fraCommentaire.Visible = False

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumElec", "cmdOKCommentaire_Click", Err, Erl
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

45      woups "frmProjSoumElec", "cmdOKDateRequise_Click", Err, Erl
End Sub

Private Sub cmdAnnulerFRS_Click()

5       On Error GoTo AfficherErreur

10      frafournisseur.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElec", "cmdAnnulerFRS_Click", Err, Erl
End Sub

Private Sub cmdAnnulerPieceTrouve_Click()

5       On Error GoTo AfficherErreur

10      fraPieceTrouve.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElec", "cmdAnnulerPieceTrouve", Err, Erl
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

165     woups "frmProjSoumElec", "cmdBonCommande_Click", Err, Erl
End Sub

Public Sub Commande()

5       On Error GoTo AfficherErreur
        
        'Change la valeur du champs "Commandé" pour la pièce à true
10      Dim rstProjet       As ADODB.Recordset
15      Dim rstPiece        As ADODB.Recordset
20      Dim rstBCPiece      As ADODB.Recordset
25      Dim rstBC           As ADODB.Recordset
30      Dim rstFRS          As ADODB.Recordset
35      Dim iIDFRS          As Integer
40      Dim sFRS            As String
45      Dim sNoBC           As String
50      Dim sWherePiece     As String
55      Dim sWhereNoLigne   As String
60      Dim sWhere          As String
65      Dim sDateRequise    As String
70      Dim sNoLigne        As String
75      Dim bPremier        As Boolean
80      Dim bPremierNoLigne As Boolean

85      Set rstProjet = New ADODB.Recordset

90      Call rstProjet.Open("SELECT ProchaineCommande FROM GRB_ProjetElec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

95      If Not IsNull(rstProjet.Fields("ProchaineCommande")) Then
100       rstProjet.Fields("ProchaineCommande") = rstProjet.Fields("ProchaineCommande") + 1

105       Call rstProjet.Update
110     End If

115     Call rstProjet.Close
120     Set rstProjet = Nothing

125     sFRS = DR_Commande.Sections("Section2").Controls("lblFournisseur").Caption
130     sNoBC = DR_Commande.Sections("Section2").Controls("lblNoBC").Caption

135     Set rstBC = New ADODB.Recordset
140     Set rstFRS = New ADODB.Recordset
145     Set rstPiece = New ADODB.Recordset
150     Set rstBCPiece = New ADODB.Recordset

155     Call rstBC.Open("SELECT * FROM GRB_BonsCommandes WHERE NoBonCommande = '" & sNoBC & "'", g_connData, adOpenDynamic, adLockOptimistic)
        
160     Do While Not rstBC.EOF
165       Call rstFRS.Open("SELECT IDFRS, NomFournisseur FROM GRB_Fournisseur WHERE IDFRS = " & rstBC.Fields("NoFournisseur"), g_connData, adOpenDynamic, adLockOptimistic)

170       If rstFRS.Fields("NomFournisseur") = sFRS Then
175         iIDFRS = rstFRS.Fields("IDFRS")

180         sDateRequise = rstBC.Fields("DateRequise")

185         Call rstFRS.Close

190         Exit Do
195       End If

200       Call rstFRS.Close

205       Call rstBC.MoveNext
210     Loop

215     Call rstBC.Close
220     Set rstBC = Nothing

225     Set rstFRS = Nothing
        
        'Ouverture du recordset du Bon de commande pour savoir quelles pièces
        'ont été commandées
230     Call rstBCPiece.Open("SELECT NoItem, NuméroLigne FROM GRB_BonsCommandes_Pieces WHERE NoFournisseur = " & iIDFRS & " AND NoBonCommande = '" & sNoBC & "'", g_connData, adOpenDynamic, adLockOptimistic)
        
        'Tant que ce n'est pas la fin des enregistrements
235     sWhere = "(IDProjet = '" & txtNoProjSoum.Text & "')"
        
240     sWherePiece = "NumItem In ("
245     sWhereNoLigne = "NuméroLigne In ("
        
250     bPremier = True
        
255     Do While Not rstBCPiece.EOF
260       If Not IsNull(rstBCPiece.Fields("NoItem")) Then
265         sNoLigne = rstBCPiece.Fields("NuméroLigne")

270         If bPremier = True Then
275           If InStr(1, sNoLigne, ",") = 0 Then
280             sWherePiece = sWherePiece & "'" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
285             sWhereNoLigne = sWhereNoLigne & rstBCPiece.Fields("NuméroLigne")
290           Else
295             bPremierNoLigne = True

300             Do While InStr(1, sNoLigne, ",") > 0
305               If bPremierNoLigne = True Then
310                 sWherePiece = sWherePiece & "'" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
315                 sWhereNoLigne = sWhereNoLigne & Left$(sNoLigne, InStr(1, sNoLigne, ",") - 1)

320                 bPremierNoLigne = False
325               Else
330                 sWherePiece = sWherePiece & ", '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
335                 sWhereNoLigne = sWhereNoLigne & ", " & Left$(sNoLigne, InStr(1, sNoLigne, ",") - 1)
340               End If

345               sNoLigne = Right$(sNoLigne, Len(sNoLigne) - (InStr(1, sNoLigne, ",") + 1))
350             Loop

355             If Trim$(sNoLigne) <> "" Then
360               sWherePiece = sWherePiece & ", '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
365               sWhereNoLigne = sWhereNoLigne & ", " & sNoLigne
370             End If
375           End If

380           bPremier = False
385         Else
390           If InStr(1, sNoLigne, ",") = 0 Then
395             sWherePiece = sWherePiece & ", '" & rstBCPiece.Fields("NoItem") & "'"
400             sWhereNoLigne = sWhereNoLigne & ", " & rstBCPiece.Fields("NuméroLigne")
405           Else
410             Do While InStr(1, sNoLigne, ",") > 0
415               sWherePiece = sWherePiece & ", '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
420               sWhereNoLigne = sWhereNoLigne & ", " & Left$(sNoLigne, InStr(1, sNoLigne, ",") - 1)

425               sNoLigne = Right$(sNoLigne, Len(sNoLigne) - (InStr(1, sNoLigne, ",") + 1))
430             Loop

435             If Trim$(sNoLigne) <> "" Then
440               sWherePiece = sWherePiece & ", '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
445               sWhereNoLigne = sWhereNoLigne & ", " & sNoLigne
450             End If
455           End If
460         End If
465       End If
    
470       Call rstBCPiece.MoveNext
475     Loop

480     sWherePiece = sWherePiece & ")"
485     sWhereNoLigne = sWhereNoLigne & ")"

490     sWhere = sWhere & " AND " & sWherePiece & " AND " & sWhereNoLigne
  
495     Call rstBCPiece.Close
500     Set rstBCPiece = Nothing
  
505     Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
  
510     Do While Not rstPiece.EOF
515       rstPiece.Fields("Commandé") = True

520       rstPiece.Fields("DateCommande") = ConvertDate(Date)

525       rstPiece.Fields("DateRequise") = sDateRequise

530       rstPiece.Fields("NomCommande") = g_sEmploye

535       rstPiece.Fields("NoSéquentiel") = Right$(sNoBC, 3)
    
540       Call rstPiece.Update
    
545       Call rstPiece.MoveNext
550     Loop
  
555     Call rstPiece.Close
560     Set rstPiece = Nothing
  
565     Call RemplirListViewProjSoum(txtNoProjSoum.Text)

570     Exit Sub

AfficherErreur:

575     woups "frmProjSoumElec", "Commande", Err, Erl
End Sub

Private Sub cmdCatalogue_Click()

5       On Error GoTo AfficherErreur

        'Pour ouvrir le catalogue électrique
10      Screen.MousePointer = vbHourglass

15      Call FrmCatalogueElec.AfficherForm(cmbPieces.Text, "", "")

20      Screen.MousePointer = vbDefault

25      Exit Sub

AfficherErreur:

30      woups "frmProjSoumElec", "cmdCatalogue_Click", Err, Erl
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

70      woups "frmProjSoumElec", "cmdForfait_Click", Err, Erl
End Sub

Private Sub cmdMauvaisPrix_Click()

5       On Error GoTo AfficherErreur

10      Call MauvaisPrix

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElec", "cmdMauvaisPrix_Click", Err, Erl
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

90                  fraPrixPiece.Tag = lvwSoumission.SelectedItem.Index

95                  m_bMauvaisPrix = True

100                 fraPrixPiece.Visible = True

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

165     woups "frmProjSoumElec", "MauvaisPrix", Err, Erl
End Sub

Private Sub cmdMaterielInutile_Click()

5       On Error GoTo AfficherErreur

10      Call MaterielInutile

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElec", "cmdMaterielInutile_Click", Err, Erl
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

155     woups "frmProjSoumElec", "MaterielInutile", Err, Erl
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
100             If ValiderFormatElectrique(sNoProjSoum) = False Then
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
      
              'S'il n'existe pas, on l'ajoute
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
                          "-  % Imprévu" & vbNewLine & _
                          "-  $ Pages manuel", vbYesNo) = vbYes Then
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

675     woups "frmProjSoumElec", "cmdCopier_Click", Err, Erl
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
  
70      Set rstOrdre = New ADODB.Recordset
  
        'Boucle pour changer la valeur de l'ordre dans le ListItem
75      For iCompteur = 1 To lvwSoumission.ListItems.count
          'Si ce n'est pas une section
80        If lvwSoumission.ListItems(iCompteur).Tag <> vbNullString Then
85          Call rstOrdre.Open("SELECT Ordre FROM GRB_SoumProjSectionElec WHERE IDSection = " & lvwSoumission.ListItems(iCompteur).Tag, g_connData, adOpenDynamic, adLockOptimistic)
        
90          lvwSoumission.ListItems(iCompteur).ListSubItems(I_COL_SOUM_MANUFACT).Tag = rstOrdre.Fields("Ordre")
      
95          Call rstOrdre.Close
100       End If
105     Next

110     Set rstOrdre = Nothing
    
115     Set rstCount = New ADODB.Recordset
    
120     Call rstCount.Open("SELECT COUNT(IDSection) as NbreSection FROM GRB_SoumProjSectionElec", g_connData, adOpenDynamic, adLockOptimistic)

125     iNbreSection = rstCount.Fields("NbreSection")

130     Call rstCount.Close
135     Set rstCount = Nothing

        'Il faut enlever les sections car ils n'ont pas d'ordre et il ne font
        'que nuire
140     For iCompteur = 1 To lvwSoumission.ListItems.count
          'Si c'est une section
145       If lvwSoumission.ListItems(iCompteur - iSection).Tag = vbNullString Then
            'On l'enlève
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

245             Call rstSection.Open("SELECT " & sSection & " FROM GRB_SoumProjSectionElec WHERE IDSection = " & lvwSoumission.ListItems(iCompteur2 + 1).Tag, g_connData, adOpenDynamic, adLockOptimistic)

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

445           itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_NET).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_PRIX_NET).Bold

450           itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_NET).Tag = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_PRIX_NET).Tag

455           itmProjSoum.SubItems(I_COL_SOUM_TEMPS) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_TEMPS)
        
460           itmProjSoum.ListSubItems(I_COL_SOUM_TEMPS).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_TEMPS).ForeColor

465           itmProjSoum.ListSubItems(I_COL_SOUM_TEMPS).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_TEMPS).Bold

470           itmProjSoum.SubItems(I_COL_SOUM_MONTAGE) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_MONTAGE)
        
475           itmProjSoum.ListSubItems(I_COL_SOUM_MONTAGE).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_MONTAGE).ForeColor

480           itmProjSoum.ListSubItems(I_COL_SOUM_MONTAGE).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_MONTAGE).Bold

485           itmProjSoum.SubItems(I_COL_SOUM_TOTAL) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_TOTAL)
        
490           itmProjSoum.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_TOTAL).ForeColor

495           itmProjSoum.ListSubItems(I_COL_SOUM_TOTAL).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_TOTAL).Bold

500           itmProjSoum.ListSubItems(I_COL_SOUM_TOTAL).Tag = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_TOTAL).Tag

505           itmProjSoum.SubItems(I_COL_SOUM_PROFIT) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_PROFIT)
        
510           itmProjSoum.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_PROFIT).ForeColor

515           itmProjSoum.ListSubItems(I_COL_SOUM_PROFIT).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_PROFIT).Bold

520           itmProjSoum.ListSubItems(I_COL_SOUM_PROFIT).Tag = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_PROFIT).Tag

525           itmProjSoum.SubItems(I_COL_SOUM_DISTRIB) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_DISTRIB)
530           itmProjSoum.ListSubItems(I_COL_SOUM_DISTRIB).Tag = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_DISTRIB).Tag
       
535           itmProjSoum.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_DISTRIB).ForeColor

540           itmProjSoum.ListSubItems(I_COL_SOUM_DISTRIB).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_DISTRIB).Bold

545           If lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_COMMENTAIRE) = "" Then
550             itmProjSoum.SubItems(I_COL_SOUM_COMMENTAIRE) = " "
555           Else
560             itmProjSoum.SubItems(I_COL_SOUM_COMMENTAIRE) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_COMMENTAIRE)
565             itmProjSoum.ListSubItems(I_COL_SOUM_COMMENTAIRE).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_COMMENTAIRE).ForeColor
570             itmProjSoum.ListSubItems(I_COL_SOUM_COMMENTAIRE).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_COMMENTAIRE).Bold
575           End If

580           If m_eType = TYPE_PROJET Then
585             If itmProjSoum.SubItems(I_COL_SOUM_PIECE) <> "" And itmProjSoum.SubItems(I_COL_SOUM_PIECE) <> "Texte" And itmProjSoum.SubItems(I_COL_SOUM_PIECE) <> "Text" Then
590               itmProjSoum.SubItems(I_COL_SOUM_ID) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_ID)

595               itmProjSoum.ListSubItems(I_COL_SOUM_ID).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_ID).ForeColor

600               itmProjSoum.ListSubItems(I_COL_SOUM_ID).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_ID).Bold

605               itmProjSoum.SubItems(I_COL_SOUM_FACTURATION) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_FACTURATION)

610               If lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_FACTURATION) = "" Then
615                 itmProjSoum.ListSubItems(I_COL_SOUM_FACTURATION).Tag = ""
620               Else
625                 itmProjSoum.ListSubItems(I_COL_SOUM_FACTURATION).Tag = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_FACTURATION).Tag
630               End If

635               itmProjSoum.SubItems(I_COL_SOUM_DATE_COMMANDE) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_DATE_COMMANDE)

640               If itmProjSoum.SubItems(I_COL_SOUM_DATE_COMMANDE) <> "" Then
645                 itmProjSoum.ListSubItems(I_COL_SOUM_DATE_COMMANDE).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_DATE_COMMANDE).ForeColor
650               End If

655               itmProjSoum.ListSubItems(I_COL_SOUM_DATE_COMMANDE).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_DATE_COMMANDE).Bold

660               itmProjSoum.ListSubItems(I_COL_SOUM_DATE_COMMANDE).Tag = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_DATE_COMMANDE).Tag

665               itmProjSoum.SubItems(I_COL_SOUM_DATE_REQUISE) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_DATE_REQUISE)

670               itmProjSoum.ListSubItems(I_COL_SOUM_DATE_REQUISE).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_DATE_REQUISE).ForeColor

675               itmProjSoum.ListSubItems(I_COL_SOUM_DATE_REQUISE).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_DATE_REQUISE).Bold

680               itmProjSoum.ListSubItems(I_COL_SOUM_DATE_REQUISE).Tag = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_DATE_REQUISE).Tag

685               itmProjSoum.SubItems(I_COL_SOUM_NOM_COMMANDE) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_NOM_COMMANDE)

690               itmProjSoum.ListSubItems(I_COL_SOUM_NOM_COMMANDE).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_NOM_COMMANDE).ForeColor

695               itmProjSoum.ListSubItems(I_COL_SOUM_NOM_COMMANDE).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_NOM_COMMANDE).Bold

700               itmProjSoum.SubItems(I_COL_SOUM_NO_SEQUENTIEL) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_NO_SEQUENTIEL)

705               itmProjSoum.ListSubItems(I_COL_SOUM_NO_SEQUENTIEL).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_NO_SEQUENTIEL).ForeColor

710               itmProjSoum.ListSubItems(I_COL_SOUM_NO_SEQUENTIEL).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_NO_SEQUENTIEL).Bold

715               itmProjSoum.SubItems(I_COL_SOUM_PROVENANCE) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_PROVENANCE)

720               itmProjSoum.ListSubItems(I_COL_SOUM_PROVENANCE).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_PROVENANCE).ForeColor

725               itmProjSoum.ListSubItems(I_COL_SOUM_PROVENANCE).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_PROVENANCE).Bold
730             End If
735           End If

740           Call lvwSoumission.ListItems.Remove(iIndexCopie)

745           Call lvwSoumission.Refresh

750           iIndex = iIndex + 1
755         End If

760         iCompteur2 = iCompteur2 + 1
765       Loop
770     Next iCompteur

775     Set rstSection = Nothing

780     If lvwSoumission.ListItems.count > 0 Then
785       Call Deselect

790       lvwSoumission.ListItems(1).Selected = True
795     End If

800     Exit Sub

AfficherErreur:

805     woups "frmProjSoumElec", "UpdateOrdre", Err, Erl
End Sub

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
          
95            Call ValeurParDefaut(itmPiece)
          
100           If itmPiece.SubItems(I_COL_SOUM_PIECE) <> "Texte" And itmPiece.SubItems(I_COL_SOUM_PIECE) <> "Text" Then
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
                      
                      
200               If Not IsNull(rstPieceFRS.Fields("PRIX_NET")) Then
205                 If Trim(rstPieceFRS.Fields("PRIX_NET")) <> vbNullString Then
210                   If rstPieceFRS.Fields("ESCOMPTE") <> vbNullString Then
215                     itmPiece.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion(rstPieceFRS.Fields("Escompte"), MODE_POURCENT)
220                   End If
           
225                   itmPiece.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(rstPieceFRS.Fields("PRIX_NET"), MODE_ARGENT, 4)
230                 Else
235                   If Not IsNull(rstPieceFRS.Fields("PRIX_SP")) Then
240                     itmPiece.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(rstPieceFRS.Fields("PRIX_SP"), MODE_ARGENT, 4)
245                   Else
250                     itmPiece.SubItems(I_COL_SOUM_PRIX_NET) = ""
255                   End If
260                 End If
265               Else
270                 If Not IsNull(rstPieceFRS.Fields("PRIX_SP")) Then
275                   itmPiece.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(rstPieceFRS.Fields("PRIX_SP"), MODE_ARGENT, 4)
280                 Else
285                   itmPiece.SubItems(I_COL_SOUM_PRIX_NET) = ""
290                 End If
295               End If

      
300               If rstPieceFRS.Fields("DeviseMonétaire") = "USA" Then
305                 itmPiece.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(itmPiece.SubItems(I_COL_SOUM_PRIX_NET)) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
310               Else
315                 If rstPieceFRS.Fields("DeviseMonétaire") = "SPA" Then
320                   itmPiece.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(itmPiece.SubItems(I_COL_SOUM_PRIX_NET)) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
325                 Else
330                   itmPiece.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(itmPiece.SubItems(I_COL_SOUM_PRIX_NET), MODE_ARGENT, 4)
335                 End If
340               End If
       
                  'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
345               itmPiece.SubItems(I_COL_SOUM_TOTAL) = Conversion(CStr(Replace(itmPiece.Text, "*", vbNullString) * itmPiece.SubItems(I_COL_SOUM_PRIX_NET) * CSng(m_sProfit)), MODE_ARGENT)
      
                  'Pour le profit, c'est le prix total - (prix net * quantité)
350               itmPiece.SubItems(I_COL_SOUM_PROFIT) = Conversion(CStr(itmPiece.SubItems(I_COL_SOUM_TOTAL) - (itmPiece.SubItems(I_COL_SOUM_PRIX_NET) * Replace(itmPiece.Text, "*", vbNullString))), MODE_ARGENT)
      
                  'Pour garder en mémoire le prix d'origine, je le mets dans le
                  'tag de la colonne Prix Listé
355               If Trim$(itmPiece.SubItems(I_COL_SOUM_PRIX_LIST)) = vbNullString Then
360                 itmPiece.SubItems(I_COL_SOUM_PRIX_LIST) = " "
365               End If
     
370               If Not IsNull(rstPieceFRS.Fields("PRIX_NET")) Then
375                 If Trim(rstPieceFRS.Fields("PRIX_NET")) <> vbNullString Then
380                   itmPiece.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = rstPieceFRS.Fields("PRIX_LIST")
385                 Else
390                   itmPiece.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = rstPieceFRS.Fields("PRIX_SP")
395                 End If
400               Else
405                 itmPiece.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = rstPieceFRS.Fields("PRIX_SP")
410               End If
415             Else
420               Call MsgBox("Il n'y a pas de prix du fournisseur " & itmPiece.SubItems(I_COL_SOUM_DISTRIB) & " pour la pièce " & itmPiece.SubItems(I_COL_SOUM_PIECE) & " ou la pièce n'existe plus!", vbOKOnly, "Erreur")
425             End If

430             Call rstPieceFRS.Close
435           End If
440         End If
445       End If
450     Next

455     Set rstPieceFRS = Nothing

460     Exit Sub

AfficherErreur:

465     woups "frmProjSoumElec", "UpdatePieces", Err, Erl
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
90            Call MsgBox("Cette soumission est fermée!", vbOKOnly, "Erreur")
95          Else
100           Call MsgBox("Cette soumission est verrouillée!", vbOKOnly, "Erreur")
105         End If
  
110         Call rstProjSoum.Close
115         Set rstProjSoum = Nothing

120         Exit Sub
125       End If

130       Call rstProjSoum.Close

135       If VerifierSiOuvert(sUser) = False Then
            'Demande du numéro de projet
140         sNoProjet = InputBox("Quel est le numéro du projet?")

145         If Trim$(sNoProjet) <> vbNullString Then
150           Screen.MousePointer = vbHourglass

155           bNoValide = True

160           If ValiderFormatNumeroProjSoum(sNoProjet) = False Then
165             bNoValide = False
170           End If

175           If bNoValide = True Then
180             If ValiderFormatElectrique(sNoProjet) = False Then
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
      
540           Call frmChoixTransfertJob.Afficher(txtNoProjSoum.Text, "E")
      
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

745     woups "frmProjSoumElec", "cmdCreerProjet_Click", Err, Erl
End Sub

Private Function VerifierSiDejaProjet() As Boolean

5       On Error GoTo AfficherErreur

        'Méthode qui sert à vérifier si une soumission est déjà assignée à un projet
10      Dim rstProjet As ADODB.Recordset
  
15      Set rstProjet = New ADODB.Recordset
  
20      Call rstProjet.Open("SELECT * FROM GRB_ProjetElec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
    
25      If Not rstProjet.EOF Then
30        VerifierSiDejaProjet = True
35      End If
    
40      Call rstProjet.Close
45      Set rstProjet = Nothing

50      Exit Function

AfficherErreur:

55      woups "frmProjSoumElec", "VerifierSiDejaProjet", Err, Erl
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
55      Set rstProjet = New ADODB.Recordset
60      Set rstSoumPiece = New ADODB.Recordset
65      Set rstProjetPiece = New ADODB.Recordset
70      Set rstEmploye = New ADODB.Recordset
75      Set rstProjSoum = New ADODB.Recordset

        'Ouverture de la soumission
80      Call rstSoum.Open("SELECT * FROM GRB_SoumissionElec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
85      Call rstSoumPiece.Open("SELECT * FROM GRB_Soumission_Pieces WHERE IDSoumission = '" & txtNoProjSoum.Text & "' AND Type = 'E' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
    
        'Ouverture du projet
90      Call rstProjet.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)
95      Call rstProjetPiece.Open("SELECT * FROM GRB_Projet_Pieces", g_connData, adOpenDynamic, adLockOptimistic)

100     Call rstProjSoum.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)
  
105     If rstProjSoum.EOF Then
110       Call rstProjSoum.AddNew
    
115       rstProjSoum.Fields("IDProjSoum") = sNoProjet
120       rstProjSoum.Fields("NoClient") = rstSoum.Fields("IDClient")
125       rstProjSoum.Fields("DateOuverture") = ConvertDate(Date)
130       rstProjSoum.Fields("Description") = rstSoum.Fields("Description")
135       rstProjSoum.Fields("Ouvert") = True
140       rstProjSoum.Fields("Type") = "P"
    
145       Call rstProjSoum.Update
150     End If
    
155     Call rstProjSoum.Close

160     Call rstProjSoum.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

165     rstProjSoum.Fields("Ouvert") = False

170     Call rstProjSoum.Update

175     Call rstProjSoum.Close
180     Set rstProjSoum = Nothing
        
        'On l'ajoute
185     Call rstProjet.AddNew
      
190     rstProjet.Fields("IDProjet") = sNoProjet
195     rstProjet.Fields("IDSoumission") = rstSoum.Fields("IDSoumission")
200     rstProjet.Fields("IDClient") = rstSoum.Fields("IDClient")
205     rstProjet.Fields("IDContact") = rstSoum.Fields("IDContact")
210     rstProjet.Fields("Description") = rstSoum.Fields("Description")
215     rstProjet.Fields("Panneau_aire") = rstSoum.Fields("Panneau_aire")
220     rstProjet.Fields("panneau_espace") = rstSoum.Fields("panneau_espace")
225     rstProjet.Fields("NbreManuel") = rstSoum.Fields("NbreManuel")
230     rstProjet.Fields("transport") = rstSoum.Fields("transport")
235     rstProjet.Fields("csa") = rstSoum.Fields("csa")
240     rstProjet.Fields("cul") = rstSoum.Fields("cul")
245     rstProjet.Fields("cur") = rstSoum.Fields("cur")
250     rstProjet.Fields("ul") = rstSoum.Fields("ul")
255     rstProjet.Fields("ur") = rstSoum.Fields("ur")
260     rstProjet.Fields("ce") = rstSoum.Fields("ce")
265     rstProjet.Fields("Delais") = rstSoum.Fields("Delais")
270     rstProjet.Fields("Creer") = ConvertDate(Date)
275     rstProjet.Fields("CheminPhotos") = rstSoum.Fields("CheminPhotos")

280     If sLiaison <> "" Then
285       rstProjet.Fields("LiaisonChargeable") = sLiaison
290     End If

295     Call rstEmploye.Open("SELECT noEmploye FROM GRB_Employés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)
    
300     rstProjet.Fields("Creer_Par") = rstEmploye.Fields("noEmploye")
   
305     Call rstEmploye.Close
310     Set rstEmploye = Nothing
    
315     rstProjet.Fields("TempsDessin") = rstSoum.Fields("TempsDessin")
320     rstProjet.Fields("TempsFabrication") = rstSoum.Fields("TempsFabrication")
325     rstProjet.Fields("TempsAssemblage") = rstSoum.Fields("TempsAssemblage")
330     rstProjet.Fields("TempsProgInterface") = rstSoum.Fields("TempsProgInterface")
335     rstProjet.Fields("TempsProgAutomate") = rstSoum.Fields("TempsProgAutomate")
340     rstProjet.Fields("TempsProgRobot") = rstSoum.Fields("TempsProgRobot")
345     rstProjet.Fields("TempsVision") = rstSoum.Fields("TempsVision")
350     rstProjet.Fields("TempsTest") = rstSoum.Fields("TempsTest")
355     rstProjet.Fields("TempsInstallation") = 0
360     rstProjet.Fields("TempsMiseService") = 0
365     rstProjet.Fields("TempsFormation") = rstSoum.Fields("TempsFormation")
370     rstProjet.Fields("TempsGestion") = rstSoum.Fields("TempsGestion")
375     rstProjet.Fields("TempsShipping") = rstSoum.Fields("TempsShipping")

380     Set rstConfig = New ADODB.Recordset

385     If Not IsNull(rstSoum.Fields("TauxDessin")) Then
390       rstProjet.Fields("TauxDessin") = rstSoum.Fields("TauxDessin")
395     Else
400       Call rstConfig.Open("SELECT TauxDessinElec FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

405       rstProjet.Fields("TauxDessin") = rstConfig.Fields("TauxDessinElec")

410       Call rstConfig.Close
415     End If

420     If Not IsNull(rstSoum.Fields("TauxFabrication")) Then
425       rstProjet.Fields("TauxFabrication") = rstSoum.Fields("TauxFabrication")
430     Else
435       Call rstConfig.Open("SELECT TauxFabrication FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

440       rstProjet.Fields("TauxFabrication") = rstConfig.Fields("TauxFabrication")

445       Call rstConfig.Close
450     End If

455     If Not IsNull(rstSoum.Fields("TauxAssemblage")) Then
460       rstProjet.Fields("TauxAssemblage") = rstSoum.Fields("TauxAssemblage")
465     Else
470       Call rstConfig.Open("SELECT TauxAssemblageElec FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

475       rstProjet.Fields("TauxAssemblage") = rstConfig.Fields("TauxAssemblageElec")

480       Call rstConfig.Close
485     End If

490     If Not IsNull(rstSoum.Fields("TauxProgInterface")) Then
495       rstProjet.Fields("TauxProgInterface") = rstSoum.Fields("TauxProgInterface")
500     Else
505       Call rstConfig.Open("SELECT TauxProgInterface FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

510       rstProjet.Fields("TauxProgInterface") = rstConfig.Fields("TauxProgInterface")

515       Call rstConfig.Close
520     End If

525     If Not IsNull(rstSoum.Fields("TauxProgAutomate")) Then
530       rstProjet.Fields("TauxProgAutomate") = rstSoum.Fields("TauxProgAutomate")
535     Else
540       Call rstConfig.Open("SELECT TauxProgAutomate FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

545       rstProjet.Fields("TauxProgAutomate") = rstConfig.Fields("TauxProgAutomate")

550       Call rstConfig.Close
555     End If

560     If Not IsNull(rstSoum.Fields("TauxProgRobot")) Then
565       rstProjet.Fields("TauxProgRobot") = rstSoum.Fields("TauxProgRobot")
570     Else
575       Call rstConfig.Open("SELECT TauxProgRobot FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

580       rstProjet.Fields("TauxProgRobot") = rstConfig.Fields("TauxProgRobot")

585       Call rstConfig.Close
590     End If

595     If Not IsNull(rstSoum.Fields("TauxVision")) Then
600       rstProjet.Fields("TauxVision") = rstSoum.Fields("TauxVision")
605     Else
610       Call rstConfig.Open("SELECT TauxVision FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

615       rstProjet.Fields("TauxVision") = rstConfig.Fields("TauxVision")

620       Call rstConfig.Close
625     End If

630     If Not IsNull(rstSoum.Fields("TauxTest")) Then
635       rstProjet.Fields("TauxTest") = rstSoum.Fields("TauxTest")
640     Else
645       Call rstConfig.Open("SELECT TauxTestElec FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

650       rstProjet.Fields("TauxTest") = rstConfig.Fields("TauxTestElec")

655       Call rstConfig.Close
660     End If

665     If Not IsNull(rstSoum.Fields("TauxInstallation")) Then
670       rstProjet.Fields("TauxInstallation") = rstSoum.Fields("TauxInstallation")
675     Else
680       Call rstConfig.Open("SELECT TauxInstallationElec FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

685       rstProjet.Fields("TauxInstallation") = rstConfig.Fields("TauxInstallationElec")

690       Call rstConfig.Close
695     End If

700     If Not IsNull(rstSoum.Fields("TauxMiseService")) Then
705       rstProjet.Fields("TauxMiseService") = rstSoum.Fields("TauxMiseService")
710     Else
715       Call rstConfig.Open("SELECT TauxMiseService FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

720       rstProjet.Fields("TauxMiseService") = rstConfig.Fields("TauxMiseService")

725       Call rstConfig.Close
730     End If

735     If Not IsNull(rstSoum.Fields("TauxFormation")) Then
740       rstProjet.Fields("TauxFormation") = rstSoum.Fields("TauxFormation")
745     Else
750       Call rstConfig.Open("SELECT TauxFormationElec FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

755       rstProjet.Fields("TauxFormation") = rstConfig.Fields("TauxFormationElec")

760       Call rstConfig.Close
765     End If

770     If Not IsNull(rstSoum.Fields("TauxGestion")) Then
775       rstProjet.Fields("TauxGestion") = rstSoum.Fields("TauxGestion")
780     Else
785       Call rstConfig.Open("SELECT TauxGestionProjetsElec FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

790       rstProjet.Fields("TauxGestion") = rstConfig.Fields("TauxGestionProjetsElec")

795       Call rstConfig.Close
800     End If

805     If Not IsNull(rstSoum.Fields("TauxShipping")) Then
810       rstProjet.Fields("TauxShipping") = rstSoum.Fields("TauxShipping")
815     Else
820       Call rstConfig.Open("SELECT TauxShippingElec FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

825       rstProjet.Fields("TauxShipping") = rstConfig.Fields("TauxShippingElec")

830       Call rstConfig.Close
835     End If

840     Set rstConfig = Nothing

845     rstProjet.Fields("PrixEmballage") = rstSoum.Fields("PrixEmballage")

850     rstProjet.Fields("imprevue") = rstSoum.Fields("imprevue")
855     rstProjet.Fields("commission") = rstSoum.Fields("commission")
860     rstProjet.Fields("Profit") = rstSoum.Fields("Profit")
865     rstProjet.Fields("total_manuel") = rstSoum.Fields("total_manuel")
870     rstProjet.Fields("total_commission") = rstSoum.Fields("total_Commission")
875     rstProjet.Fields("total_profit") = rstSoum.Fields("Total_Profit")
880     rstProjet.Fields("PrixTotal") = rstSoum.Fields("PrixTotal")
885     rstProjet.Fields("total_piece") = rstSoum.Fields("Total_piece")
890     rstProjet.Fields("total_imprevue") = rstSoum.Fields("total_imprevue")
895     rstProjet.Fields("total_temps") = rstSoum.Fields("total_temps")
900     rstProjet.Fields("SansTemps") = rstSoum.Fields("SansTemps")
905     rstProjet.Fields("MontantForfait") = rstSoum.Fields("MontantForfait")
910     rstProjet.Fields("InitialeForfait") = rstSoum.Fields("InitialeForfait")
915     rstProjet.Fields("ProchaineCommande") = 1

920     Call rstProjet.Update
    
        'Ajout des pièces
925     Do While Not rstSoumPiece.EOF
930       If rstSoumPiece.Fields("TransfertJob") = True Then
935         Call rstProjetPiece.AddNew

940         rstProjetPiece.Fields("Type") = "E"
  
945         rstProjetPiece.Fields("IDProjet") = sNoProjet
950         rstProjetPiece.Fields("IDSection") = rstSoumPiece.Fields("IDSection")
955         rstProjetPiece.Fields("NumItem") = rstSoumPiece.Fields("NumItem")
960         rstProjetPiece.Fields("Qté") = rstSoumPiece.Fields("Qté")
965         rstProjetPiece.Fields("Desc_FR") = rstSoumPiece.Fields("Desc_FR")
970         rstProjetPiece.Fields("Desc_EN") = rstSoumPiece.Fields("Desc_EN")
975         rstProjetPiece.Fields("Manufact") = rstSoumPiece.Fields("Manufact")
980         rstProjetPiece.Fields("Prix_List") = rstSoumPiece.Fields("Prix_list")
985         rstProjetPiece.Fields("Escompte") = rstSoumPiece.Fields("Escompte")
990         rstProjetPiece.Fields("Prix_net") = rstSoumPiece.Fields("Prix_net")
995         rstProjetPiece.Fields("OrdreSection") = rstSoumPiece.Fields("OrdreSection")
1000        rstProjetPiece.Fields("NuméroLigne") = rstSoumPiece.Fields("NuméroLigne")
1005        rstProjetPiece.Fields("IDFRS") = rstSoumPiece.Fields("IDFRS")
1010        rstProjetPiece.Fields("Temps") = rstSoumPiece.Fields("Temps")
1015        rstProjetPiece.Fields("Temps_total") = rstSoumPiece.Fields("Temps_Total")
1020        rstProjetPiece.Fields("Prix_total") = rstSoumPiece.Fields("Prix_Total")
1025        rstProjetPiece.Fields("Profit_argent") = rstSoumPiece.Fields("Profit_argent")
1030        rstProjetPiece.Fields("SousSection") = rstSoumPiece.Fields("SousSection")
1035        rstProjetPiece.Fields("PrixOrigine") = rstSoumPiece.Fields("PrixOrigine")
1040        rstProjetPiece.Fields("Visible") = rstSoumPiece.Fields("Visible")
1045        rstProjetPiece.Fields("Commentaire") = rstSoumPiece.Fields("Commentaire")
1050        rstProjetPiece.Fields("Quoté") = rstSoumPiece.Fields("Quoté")
1055        rstProjetPiece.Fields("Devise") = rstSoumPiece.Fields("Devise")

1060        Call rstProjetPiece.Update

1065        If sLiaison <> "" Then
1070          If Right$(sNoProjet, 2) >= "60" And Right$(sNoProjet, 2) <= 79 Then
1075
1080          Else
1085            If Right$(sNoProjet, 2) >= 80 And Right$(sNoProjet, 2) <= 98 Then
1090
1095            End If
1100          End If
1105        End If
1110      End If
   
1115      Call rstSoumPiece.MoveNext
1120    Loop

1125    m_eType = TYPE_PROJET

1130    If CDbl(rstSoum.Fields("TempsInstallation")) > 0 Or CDbl(rstSoum.Fields("TempsMiseService")) > 0 Then
1135      Call CreerProjetInstallation(Left$(sNoProjet, 7) & "51")
1140    End If

1145    Call rstSoum.Close
1150    Set rstSoum = Nothing

1155    Call rstProjet.Close
1160    Set rstProjet = Nothing

1165    Call rstSoumPiece.Close
1170    Set rstSoumPiece = Nothing
  
1175    Call rstProjetPiece.Close
1180    Set rstProjetPiece = Nothing

1185    Call CalculerTotalRecordset(sNoProjet)

1190    Exit Sub

AfficherErreur:

1195    woups "frmProjSoumElec", "TransfererSoumDansProjet", Err, Erl
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
45      Set rstProjet = New ADODB.Recordset
50      Set rstEmploye = New ADODB.Recordset
55      Set rstProjSoum = New ADODB.Recordset

        'Ouverture de la soumission
60      Call rstSoum.Open("SELECT * FROM GRB_SoumissionElec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
    
        'Ouverture du projet
65      Call rstProjet.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)

70      If rstProjet.EOF Then
75        Call rstProjSoum.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)
  
80        If rstProjSoum.EOF Then
85          Call rstProjSoum.AddNew
    
90          rstProjSoum.Fields("IDProjSoum") = sNoProjet
95          rstProjSoum.Fields("NoClient") = rstSoum.Fields("IDClient")
100         rstProjSoum.Fields("Description") = rstSoum.Fields("Description")
105         rstProjSoum.Fields("DateOuverture") = ConvertDate(Date)
110         rstProjSoum.Fields("Ouvert") = True
115         rstProjSoum.Fields("Type") = "P"
    
120         Call rstProjSoum.Update
125       End If
    
130       Call rstProjSoum.Close
135       Set rstProjSoum = Nothing
        
          'On l'ajoute
140       Call rstProjet.AddNew
      
145       rstProjet.Fields("IDProjet") = sNoProjet
150       rstProjet.Fields("IDSoumission") = vbNullString
155       rstProjet.Fields("IDClient") = rstSoum.Fields("IDClient")
160       rstProjet.Fields("IDContact") = rstSoum.Fields("IDContact")
165       rstProjet.Fields("Description") = rstSoum.Fields("Description")
170       rstProjet.Fields("Panneau_aire") = rstSoum.Fields("Panneau_aire")
175       rstProjet.Fields("panneau_espace") = rstSoum.Fields("panneau_espace")
180       rstProjet.Fields("NbreManuel") = rstSoum.Fields("NbreManuel")
185       rstProjet.Fields("transport") = rstSoum.Fields("transport")
190       rstProjet.Fields("csa") = rstSoum.Fields("csa")
195       rstProjet.Fields("cul") = rstSoum.Fields("cul")
200       rstProjet.Fields("cur") = rstSoum.Fields("cur")
205       rstProjet.Fields("ul") = rstSoum.Fields("ul")
210       rstProjet.Fields("ur") = rstSoum.Fields("ur")
215       rstProjet.Fields("ce") = rstSoum.Fields("ce")
220       rstProjet.Fields("Delais") = rstSoum.Fields("Delais")
225       rstProjet.Fields("Creer") = ConvertDate(Date)
230       rstProjet.Fields("CheminPhotos") = rstSoum.Fields("CheminPhotos")
  
235       Call rstEmploye.Open("SELECT noEmploye FROM GRB_Employés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)
    
240       rstProjet.Fields("Creer_Par") = rstEmploye.Fields("noEmploye")
   
245       Call rstEmploye.Close
250       Set rstEmploye = Nothing
    
255       rstProjet.Fields("TempsDessin") = 0
260       rstProjet.Fields("TempsFabrication") = 0
265       rstProjet.Fields("TempsAssemblage") = 0
270       rstProjet.Fields("TempsProgInterface") = 0
275       rstProjet.Fields("TempsProgAutomate") = 0
280       rstProjet.Fields("TempsProgRobot") = 0
285       rstProjet.Fields("TempsVision") = 0
290       rstProjet.Fields("TempsTest") = 0
295       rstProjet.Fields("TempsInstallation") = rstSoum.Fields("TempsInstallation")
300       rstProjet.Fields("TempsMiseService") = rstSoum.Fields("TempsMiseService")
305       rstProjet.Fields("TempsFormation") = 0
310       rstProjet.Fields("TempsGestion") = 0
315       rstProjet.Fields("TempsShipping") = 0

320       Set rstConfig = New ADODB.Recordset

325       If Not IsNull(rstSoum.Fields("TauxDessin")) Then
330         rstProjet.Fields("TauxDessin") = rstSoum.Fields("TauxDessin")
335       Else
340         Call rstConfig.Open("SELECT TauxDessinElec FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

345         rstProjet.Fields("TauxDessin") = rstConfig.Fields("TauxDessinElec")

350         Call rstConfig.Close
355       End If

360       If Not IsNull(rstSoum.Fields("TauxFabrication")) Then
365         rstProjet.Fields("TauxFabrication") = rstSoum.Fields("TauxFabrication")
370       Else
375         Call rstConfig.Open("SELECT TauxFabrication FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

380         rstProjet.Fields("TauxFabrication") = rstConfig.Fields("TauxFabrication")

385         Call rstConfig.Close
390       End If

395       If Not IsNull(rstSoum.Fields("TauxAssemblage")) Then
400         rstProjet.Fields("TauxAssemblage") = rstSoum.Fields("TauxAssemblage")
405       Else
410         Call rstConfig.Open("SELECT TauxAssemblageElec FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

415         rstProjet.Fields("TauxAssemblage") = rstConfig.Fields("TauxAssemblageElec")

420         Call rstConfig.Close
425       End If

430       If Not IsNull(rstSoum.Fields("TauxProgInterface")) Then
435         rstProjet.Fields("TauxProgInterface") = rstSoum.Fields("TauxProgInterface")
440       Else
445         Call rstConfig.Open("SELECT TauxProgInterface FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

450         rstProjet.Fields("TauxProgInterface") = rstConfig.Fields("TauxProgInterface")

455         Call rstConfig.Close
460       End If

465       If Not IsNull(rstSoum.Fields("TauxProgAutomate")) Then
470         rstProjet.Fields("TauxProgAutomate") = rstSoum.Fields("TauxProgAutomate")
475       Else
480         Call rstConfig.Open("SELECT TauxProgAutomate FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

485         rstProjet.Fields("TauxProgAutomate") = rstConfig.Fields("TauxProgAutomate")

490         Call rstConfig.Close
495       End If

500       If Not IsNull(rstSoum.Fields("TauxProgRobot")) Then
505         rstProjet.Fields("TauxProgRobot") = rstSoum.Fields("TauxProgRobot")
510       Else
515         Call rstConfig.Open("SELECT TauxProgRobot FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

520         rstProjet.Fields("TauxProgRobot") = rstConfig.Fields("TauxProgRobot")

525         Call rstConfig.Close
530       End If

535       If Not IsNull(rstSoum.Fields("TauxVision")) Then
540         rstProjet.Fields("TauxVision") = rstSoum.Fields("TauxVision")
545       Else
550         Call rstConfig.Open("SELECT TauxVision FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

555         rstProjet.Fields("TauxVision") = rstConfig.Fields("TauxVision")

560         Call rstConfig.Close
565       End If

570       If Not IsNull(rstSoum.Fields("TauxTest")) Then
575         rstProjet.Fields("TauxTest") = rstSoum.Fields("TauxTest")
580       Else
585         Call rstConfig.Open("SELECT TauxTestElec FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

590         rstProjet.Fields("TauxTest") = rstConfig.Fields("TauxTestElec")

595         Call rstConfig.Close
600       End If

605       If Not IsNull(rstSoum.Fields("TauxInstallation")) Then
610         rstProjet.Fields("TauxInstallation") = rstSoum.Fields("TauxInstallation")
615       Else
620         Call rstConfig.Open("SELECT TauxInstallationElec FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

625         rstProjet.Fields("TauxInstallation") = rstConfig.Fields("TauxInstallationElec")

630         Call rstConfig.Close
635       End If

640       If Not IsNull(rstSoum.Fields("TauxMiseService")) Then
645         rstProjet.Fields("TauxMiseService") = rstSoum.Fields("TauxMiseService")
650       Else
655         Call rstConfig.Open("SELECT TauxMiseService FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

660         rstProjet.Fields("TauxMiseService") = rstConfig.Fields("TauxMiseService")

665         Call rstConfig.Close
670       End If

675       If Not IsNull(rstSoum.Fields("TauxFormation")) Then
680         rstProjet.Fields("TauxFormation") = rstSoum.Fields("TauxFormation")
685       Else
690         Call rstConfig.Open("SELECT TauxFormationElec FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

695         rstProjet.Fields("TauxFormation") = rstConfig.Fields("TauxFormationElec")

700         Call rstConfig.Close
705       End If

710       If Not IsNull(rstSoum.Fields("TauxGestion")) Then
715         rstProjet.Fields("TauxGestion") = rstSoum.Fields("TauxGestion")
720       Else
725         Call rstConfig.Open("SELECT TauxGestionProjetsElec FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

730         rstProjet.Fields("TauxGestion") = rstConfig.Fields("TauxGestionProjetsElec")

735         Call rstConfig.Close
740       End If

745       If Not IsNull(rstSoum.Fields("TauxShipping")) Then
750         rstProjet.Fields("TauxShipping") = rstSoum.Fields("TauxShipping")
755       Else
760         Call rstConfig.Open("SELECT TauxShippingElec FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

765         rstProjet.Fields("TauxShipping") = rstConfig.Fields("TauxShippingElec")

770         Call rstConfig.Close
775       End If

780       Set rstConfig = Nothing

785       rstProjet.Fields("PrixEmballage") = 0

790       rstProjet.Fields("imprevue") = rstSoum.Fields("imprevue")
795       rstProjet.Fields("commission") = rstSoum.Fields("commission")
800       rstProjet.Fields("Profit") = rstSoum.Fields("Profit")
805       rstProjet.Fields("total_manuel") = rstSoum.Fields("total_manuel")
810       rstProjet.Fields("total_commission") = rstSoum.Fields("total_Commission")
815       rstProjet.Fields("total_profit") = rstSoum.Fields("Total_Profit")
820       rstProjet.Fields("PrixTotal") = rstSoum.Fields("PrixTotal")
825       rstProjet.Fields("total_piece") = rstSoum.Fields("Total_piece")
830       rstProjet.Fields("total_imprevue") = rstSoum.Fields("total_imprevue")
835       rstProjet.Fields("total_temps") = rstSoum.Fields("total_temps")
840       rstProjet.Fields("SansTemps") = rstSoum.Fields("SansTemps")
845       rstProjet.Fields("ProchaineCommande") = 1

850       Call rstProjet.Update
855     End If
    
860     Call rstSoum.Close
865     Set rstSoum = Nothing

870     Call rstProjet.Close
875     Set rstProjet = Nothing

880     Call CalculerTotalRecordset(sNoProjet)

885     Exit Sub

AfficherErreur:

890     woups "frmProjSoumElec", "CreerProjetInstallation", Err, Erl
End Sub

Private Sub cmdDate_Click()

5       On Error GoTo AfficherErreur

        'Ouverture du calendrier
10      If Trim$(txtDelais.Text) <> vbNullString Then
15        mvwDate.Value = txtDelais.Text
20      Else
25        mvwDate.Value = Date
30      End If
  
35      mvwDate.Visible = True
  
40      Call mvwDate.SetFocus

45      Exit Sub

AfficherErreur:

50      woups "frmProjSoumElec", "cmdDate_Click", Err, Erl
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

55      woups "frmProjSoumElec", "cmdDateFacturation_Click", Err, Erl
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
60          Call frmChoixDemande.AfficherProjetSoumission(txtNoProjSoum.Text, ELECTRIQUE, MODE_PIECE, m_eType)
65        Else
70          If m_eType = TYPE_PROJET Then
75            Call MsgBox("Ce projet est en modification par " & sUser & "!", vbOKOnly, "Erreur")
80          Else
85            Call MsgBox("Cette soumission est en modification par " & sUser & "!", vbOKOnly, "Erreur")
90          End If
95        End If
100      Else
105        If rstProjSoum.Fields("Ouvert") = False Then
110          If m_eType = TYPE_PROJET Then
115            Call MsgBox("Ce projet est fermé!", vbOKOnly, "Erreur")
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

190     woups "frmProjSoumElec", "cmdDemande_Click", Err, Erl
End Sub

Private Sub cmdExtra_Click()

5       On Error GoTo AfficherErreur

10      Dim sNumero     As String
15      Dim rstProjSoum As ADODB.Recordset
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
125         If ValiderFormatElectrique(sNumero) = False Then
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
220         bExiste = True

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
            
            'On recalcul le prix
495         Call CalculerPrix
  
            'Débarre les champs
500         Call BarrerChamps(False)
            
505         m_sTempsDessin = "0"
510         m_sTempsFabrication = "0"
515         m_sTempsAssemblage = "0"
520         m_sTempsProgInterface = "0"
525         m_sTempsProgAutomate = "0"
530         m_sTempsProgRobot = "0"
535         m_sTempsVision = "0"
540         m_sTempsTest = "0"
545         m_sTempsInstallation = "0"
550         m_sTempsMiseService = "0"
555         m_sTempsFormation = "0"
560         m_sTempsGestion = "0"
565         m_sTempsShipping = "0"
                        
570         m_sNbrePersonne = "0"
575         m_sTempsHebergement = "0"
580         m_sTempsRepas = "0"
585         m_sTempsTransport = "0"
590         m_sTempsUniteMobile = "0"
595         m_sPrixEmballage = "0"
                        
600         txtNbreManuel.Text = "0"
605         txtPrixManuel.Text = "0"

610         txtForfait.Text = ""
615         lblForfaitInitiale.Caption = ""

620         txtPrixReception.Text = "0"
625         txtPrixSoumission.Text = "0"
      
630         txtPrixTotal.Text = "0"
635         txtProfit.Text = "0"
640         txtCommission.Text = "0"
645         txtTotalTemps.Text = "0"
650         txtTotalPieces.Text = "0"
655         txtImprevus.Text = "0"
660         txtNoSoumission.Text = vbNullString
      
            'Vide la valeur par défaut si demande Sous-Section
665         m_sSousSection = vbNullString

670         txtProjet.Text = vbNullString
        
675         m_bModeAjout = True
680         m_bModeAffichage = False
685         m_bExtra = True
                
690         lvwSoumission.Height = lvwSoumission.Height * 0.49
695         lvwSoumission.Top = lvwPieces.Top + lvwPieces.Height + (lvwSoumission.Height * 0.02)
        
            'Met le form en mode ajout/modif
700         Call AfficherControles(MODE_AJOUT_MODIF)
705       End If
710     End If
  
715     Screen.MousePointer = vbDefault

720     Exit Sub

AfficherErreur:

725     woups "frmProjSoumElec", "cmdExtra_Click", Err, Erl
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

40      woups "frmProjSoumElec", "cmdHistorique_Click", Err, Erl
End Sub

Private Sub cmdBavards_Click()

5       On Error GoTo AfficherErreur

        'Ouverture de l'historique des suppressions de pièces
10      If cmbProjSoum.ListCount > 0 Then
15        Call RemplirListViewSuppression

20        lvwBavard.Visible = True
  
25        Call lvwBavard.SetFocus
30      End If

35      Exit Sub

AfficherErreur:

40      woups "frmProjSoumElec", "cmdBavards_Click", Err, Erl
End Sub

Private Sub cmdLegende_Click()

5       On Error GoTo AfficherErreur

10      Call OuvrirForm(frmLegende, True)

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElec", "cmdLegende_Click", Err, Erl
End Sub

Private Sub cmdOKDateRequise_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
5       On Error GoTo AfficherErreur

10      If m_bMonthViewHasFocus = True Then
15        Call cmdOKDateRequise_Click
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmProjSoumElec", "cmdOKDateRequise_MouseUp", Err, Erl
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

40      woups "frmProjSoumElec", "cmdOKFRS_Click", Err, Erl
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

70      woups "frmProjSoumElec", "cmdOKPieceTrouve_Click", Err, Erl
End Sub

Private Sub cmdRafraichir_Click()

5       On Error GoTo AfficherErreur

10      If m_sTri <> vbNullString Then
15        m_sTri = vbNullString
  
20        Call RemplirListViewPieces
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmProjSoumElec", "cmdRafraichir_Click", Err, Erl
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
50              Set rstProjSoum = New ADODB.Recordset

                'Ouvre les tables
55              Call rstProjSoum.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & txtNoProjSoum.Text & "' ORDER BY IDProjet", g_connData, adOpenDynamic, adLockOptimistic)

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

140     woups "frmProjSoumElec", "cmdImprimer_Click", Err, Erl
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

40      woups "frmProjSoumElec", "cmdRechercherClient_Click", Err, Erl
End Sub

Private Sub cmdReset_Click()
        'Permet d'effacer le champs Modification et Par si c'est le user actuel
5       On Error GoTo AfficherErreur

10      Dim rstProjSoum As ADODB.Recordset

15      If MsgBox("Êtes-vous certains de ne pas être en modification sur un autre ordinateur?", vbYesNo) = vbYes Then
20        Set rstProjSoum = New ADODB.Recordset

25        If m_eType = TYPE_PROJET Then
30          Call rstProjSoum.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
35        Else
40          Call rstProjSoum.Open("SELECT * FROM GRB_SoumissionElec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
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

90      woups "frmProjSoumElec", "cmdReset_Click", Err, Erl
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

55      woups "frmProjSoumElec", "cmdRetour_Click", Err, Erl
End Sub

Private Sub cmdSortieMagasin_Click()

5       On Error GoTo AfficherErreur

10      Call SortieMagasin

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElec", "cmdSortieMagasin_Click", Err, Erl
End Sub

Private Sub ChangerQuantite()

5       On Error GoTo AfficherErreur

10      Dim sQuantite As String
15      Dim itmSoum   As ListItem

20      sQuantite = InputBox("Quelle est la nouvelle quantité?")

25      If IsNumeric(sQuantite) Then
30        Set itmSoum = lvwSoumission.SelectedItem

35        itmSoum.Text = sQuantite

40        If Trim$(itmSoum.SubItems(I_COL_SOUM_TEMPS)) <> vbNullString Then
            'On calcul le temps * quantité pour la colonne montage
45          itmSoum.SubItems(I_COL_SOUM_MONTAGE) = CDbl(Replace(itmSoum.SubItems(I_COL_SOUM_TEMPS), ".", ",")) * CDbl(Replace(itmSoum.Text, "*", vbNullString))
50        Else
55          itmSoum.SubItems(I_COL_SOUM_MONTAGE) = vbNullString
60        End If
      
          'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
65        itmSoum.SubItems(I_COL_SOUM_TOTAL) = Conversion(Round(Replace(itmSoum.Text, "*", vbNullString) * itmSoum.SubItems(I_COL_SOUM_PRIX_NET) * CSng(m_sProfit), 2), MODE_ARGENT, 2)
      
          'Pour le profit, c'est le prix total - (prix net * quantité)
70        itmSoum.SubItems(I_COL_SOUM_PROFIT) = Conversion(CStr(Round(itmSoum.SubItems(I_COL_SOUM_TOTAL) - (itmSoum.SubItems(I_COL_SOUM_PRIX_NET) * Replace(itmSoum.Text, "*", vbNullString)), 2)), MODE_ARGENT)

75        Call CalculerTempsFabrication

80        Call CalculerPrix
85      End If

90      Exit Sub

AfficherErreur:

95      woups "frmProjSoumElec", "ChangerQuantite", Err, Erl
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

75                    sTag = ""
80                  End If
                
85                  lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_DESCR).ForeColor = lColor
90                  lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = lColor
95                  lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = lColor
100                 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = lColor
105                 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_MONTAGE).ForeColor = lColor
110                 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor = lColor
115                 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = lColor
120                 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = lColor
125                 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = lColor
130                 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_TEMPS).ForeColor = lColor
135                 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = lColor

140                 If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_FACTURATION) = "" Then
145                   lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_FACTURATION) = " "
150                 End If

155                 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_FACTURATION).Tag = sTag

160                 Call lvwSoumission.Refresh

165                 Call CalculerPrixReception
170               End If
175             End If
180           End If
185         End If
190       Else
195         Call MsgBox("Cette commande doit être faite dans le projet " & lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PROVENANCE), vbOKOnly, "Erreur")
200       End If
205     End If

210     Exit Sub

AfficherErreur:

215     woups "frmProjSoumElec", "SortieMagasin", Err, Erl
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
60                  If itmProjet.ListSubItems(I_COL_SOUM_FACTURATION).Tag <> "" Then
                      'On ajoute le montant
65                    dblPrixReception = Round(dblPrixReception + (itmProjet.SubItems(I_COL_SOUM_PRIX_NET) * Replace(itmProjet.ListSubItems(I_COL_SOUM_FACTURATION).Tag, "*", "")), 2)
70                  Else
75                    dblPrixReception = Round(dblPrixReception, 2)
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

155     woups "frmProjSoumElec", "CalculerPrixReception", Err, Erl
End Sub

Private Sub cmdSupprimerFRS_Click()
        'Permet d'effacer un Fournisseur
5       On Error GoTo AfficherErreur

10      Dim sPiece As String

        'Si c'est pas "Choisir ultérieurement"
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

65        If MsgBox("Voulez-vous vraiment supprimer le fournisseur " & lvwfournisseur.SelectedItem.Text & " pour la pièce " & sPiece & "?", vbYesNo, "Suppression") = vbYes Then
70          Call g_connData.Execute("DELETE * FROM GRB_PiecesFRS WHERE NoEnreg = " & lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_ENTRER_PAR).Tag)

75          Call RemplirListViewFournisseur

80          frafournisseur.Visible = True

85          Call lvwfournisseur.SetFocus
90        End If
95      End If

100     Exit Sub

AfficherErreur:

105     woups "frmProjSoumElec", "cmdSupprimerFRS_Click", Err, Erl
End Sub

Private Sub cmdTemps_Click()

5       On Error GoTo AfficherErreur

10      If cmbProjSoum.ListCount > 0 Then
15        If m_eMode = MODE_AJOUT_MODIF Then
20          If m_bModeAjout = True Then
25            If m_bExtra = True Then
30              Call frmProjSoumElecTemps.Afficher(txtNoProjSoum.Text, m_eType, m_eMode, False)
35            Else
40              Call frmProjSoumElecTemps.Afficher(txtNoProjSoum.Text, m_eType, m_eMode, True)
45            End If
50          Else
55            Call frmProjSoumElecTemps.Afficher(txtNoProjSoum.Text, m_eType, m_eMode, False)
60          End If
65        Else
70          Call frmProjSoumElecTemps.Afficher(txtNoProjSoum.Text, m_eType, m_eMode, False)
75        End If
80      End If

85      If m_eMode = MODE_AJOUT_MODIF Then
90        Call CalculerPrix
95      End If
  
100     m_bTempsDejaOuvert = True

105     Exit Sub

AfficherErreur:

110     woups "frmProjSoumElec", "cmdTemps_Click", Err, Erl
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

165     woups "frmProjSoumElec", "cmdTexte_Click", Err, Erl
End Sub

Private Sub AjouterTexte(ByVal iIndex As Integer, ByVal sTexte As String, ByVal sNomSousSection As String)

5       On Error GoTo AfficherErreur

        'Méthode pour ajouter le texte
10      Dim sSousSection As String
15      Dim sOrdre       As String
20      Dim sIDSection   As String

        'S'il faut l'ajouter à la fin, on prend les infos du dernier enregistrement
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

300     woups "frmProjSoumElec", "AjouterTexte", Err, Erl
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

40      woups "frmProjSoumElec", "cmdTri_Click", Err, Erl
End Sub

Private Sub cmdPhotos_Click()

5       On Error GoTo AfficherErreur
  
10      If txtCheminPhotos.Text <> vbNullString Then
15        Call frmPhotoProjSoum.Afficher(txtCheminPhotos.Text)
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmProjSoumElec", "cmdPhotos_Click", Err, Erl
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
45        If Forms(iCompteur).Name = "FrmReceptionElec" Then
50          bOuvert = True

55          Exit For
60        End If
65      Next

70      If bOuvert = True Then
75        Call Unload(FrmReceptionElec)
80      End If

85      Call FrmReceptionElec.AfficherProjet(g_sUserID, txtNoProjSoum.Text)

90      Call RemplirListViewProjSoum(txtNoProjSoum.Text)

95      Exit Sub

AfficherErreur:

100     woups "frmProjSoumElec", "cmdReception_Click", Err, Erl
End Sub

Private Sub cmdBrowse_Click()

5       On Error GoTo AfficherErreur

10      Call OuvrirForm(frmChoixDossier, True)

15      If m_bAnnulerChemin = False Then
20        txtCheminPhotos.Text = m_sChemin
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmProjSoumElec", "cmdBrowse_Click", Err, Erl
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

5       On Error GoTo AfficherErreur

10      If m_eMode = MODE_AJOUT_MODIF Then
15        Call MsgBox("Veuillez enregistrer ou annuler avant de fermer!", vbOKOnly, "Erreur")

20        Cancel = 1
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmProjSoumElec", "Form_QueryUnload", Err, Erl
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

170     woups "frmProjSoumElec", "Form_Resize", Err, Erl
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

180     woups "frmProjSoumElec", "PositionnerBoutons", Err, Erl
End Sub

Private Sub Form_Unload(Cancel As Integer)

5       On Error GoTo AfficherErreur

10      Set FrmProjSoumElec = Nothing

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElec", "Form_Unload", Err, Erl
End Sub

Private Sub lvwHistorique_LostFocus()

5       On Error GoTo AfficherErreur

        'Lorsque l'historique perd le focus, on l'enlève
10      lvwHistorique.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElec", "lvwHistorique_LostFocus", Err, Erl
End Sub

Private Sub lvwBavard_LostFocus()

5       On Error GoTo AfficherErreur

        'Lorsque le bavard perd le focus, on l'enlève
10      lvwBavard.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElec", "lvwBavard_LostFocus", Err, Erl
End Sub

Private Sub lvwfournisseur_KeyDown(KeyCode As Integer, Shift As Integer)
        'Permet d'effacer un Fournisseur
5       On Error GoTo AfficherErreur

10      Dim sPiece As String

15      If KeyCode = vbKeyDelete Then
20        If g_bModificationCatalogueElec = True Then
            'Si c'est pas Choisir Ultérieurement
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

125     woups "frmProjSoumElec", "lvwFournisseur_KeyDown", Err, Erl
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
50           Call MsgBox("Aucun enregistrements trouvés!", vbOKOnly, "Erreur")
55          End If
60        Else
65          Call MsgBox("Il faut un minimum de 2 caractères pour rechercher!", vbOKOnly, "Erreur")
70        End If
75      End If

80      Exit Sub

AfficherErreur:

85      woups "frmProjSoumElec", "lvwPieces_ColumnClick", Err, Erl
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
125       Call rstPiece.Open("SELECT * FROM GRB_CatalogueElec WHERE INSTR(1," & sChamps & ",'" & sRecherche & "') > 0 ", g_connData, adOpenDynamic, adLockOptimistic)
130     Else
135       Call rstPiece.Open("SELECT * FROM GRB_CatalogueElec WHERE INSTR(1," & sChamps & ",'" & Replace(sTexte, "'", "''") & "')> 0 ", g_connData, adOpenDynamic, adLockOptimistic)
140     End If

        'Pour chaque enregistrement
145     Do While Not rstPiece.EOF
          'On ajoute dans le ListView
150       Set itmPiece = lvwPieceTrouve.ListItems.Add

155       If Not IsNull(rstPiece.Fields("TEMPS")) Then
160         itmPiece.Tag = rstPiece.Fields("TEMPS")
165       Else
170         itmPiece.Tag = vbNullString
175       End If

180       If Not IsNull(rstPiece.Fields("PIECE_GRB")) Then
185         itmPiece.Text = rstPiece.Fields("PIECE_GRB")
190       Else
195         itmPiece.Text = ""
200       End If

205       itmPiece.SubItems(I_COL_RECH_NO_ITEM) = rstPiece.Fields("PIECE")
210       itmPiece.SubItems(I_COL_RECH_CATEGORIE) = rstPiece.Fields("CATEGORIE")

215       If Not IsNull(rstPiece.Fields("FABRICANT")) Then
220         itmPiece.SubItems(I_COL_RECH_MANUFACT) = rstPiece.Fields("FABRICANT")
225       Else
230         itmPiece.SubItems(I_COL_RECH_MANUFACT) = ""
235       End If

240       If Not IsNull(rstPiece.Fields("DESC_EN")) Then
245         itmPiece.SubItems(I_COL_RECH_DESCR_EN) = rstPiece.Fields("DESC_EN")
250       Else
255         itmPiece.SubItems(I_COL_RECH_DESCR_EN) = ""
260       End If

265       If Not IsNull(rstPiece.Fields("DESC_FR")) Then
270         itmPiece.SubItems(I_COL_RECH_DESCR_FR) = rstPiece.Fields("DESC_FR")
275       Else
280         itmPiece.SubItems(I_COL_RECH_DESCR_FR) = ""
285       End If

290       Call rstPiece.MoveNext
295     Loop

300     Call rstPiece.Close

305     Set rstPiece = Nothing

310     Exit Sub

AfficherErreur:

315     woups "frmProjSoumElec", "RemplirListViewRecherche", Err, Erl
End Sub

Private Sub lvwPieces_KeyDown(KeyCode As Integer, Shift As Integer)

5       On Error GoTo AfficherErreur

10      Dim sTexte As String

15      If Shift = vbCtrlMask Then
20        If KeyCode = vbKeyF Then
25          sTexte = InputBox("Quel est le texte à rechercher?")

30          If Trim$(sTexte) <> vbNullString Then
35            If Len(Trim$(sTexte)) >= 2 Then
40              Call RemplirListViewRecherche(1, sTexte)

45              If lvwPieceTrouve.ListItems.count > 0 Then
50                fraPieceTrouve.Visible = True
55              Else
60               Call MsgBox("Aucun enregistrements trouvés!", vbOKOnly, "Erreur")
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

125     woups "frmProjSoumElec", "lvwPieces_KeyDown", Err, Erl
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

80      woups "frmProjSoumElec", "lvwPieceTrouve_DblClick", Err, Erl
End Sub

Private Sub lvwSoumission_DblClick()

5       On Error GoTo AfficherErreur
        
        'Si il y a des enregistrements
10      If lvwSoumission.ListItems.count > 0 Then
15        If m_eMode = MODE_AJOUT_MODIF Then
            'Si c'est une sous-section
20          If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) = vbNullString Then
25            Call ModifierSousSection
30          Else
              'Si c'est une pièce
35            If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) <> vbNullString And lvwSoumission.SelectedItem.Tag <> vbNullString Then
                'Si la pièce n'est pas un Text ou Texte
40              If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) <> "Texte" And lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) <> "Text" Then
                  'Si la pièce n'a pas de fournisseur
45                If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_DISTRIB) = vbNullString Then
50                  If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PROFIT) = "" Then
55                    lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PROFIT) = " "
60                  End If

65                  If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PROFIT).Tag <> "EXTRA" Then
70                    Call AjouterPrix
75                  Else
80                    Call MsgBox("Cette commande doit être faite dans le projet " & lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PROVENANCE), vbOKOnly, "Erreur")
85                  End If
90                Else
                    'Si la pièce est en commande
95                  If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_ORANGE Then
100                   If MsgBox("Voulez-vous annuler cette commande?", vbYesNo) = vbYes Then
105                     Call AnnulerCommande
110                   End If
115                 Else
120                   If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_BRUN Then
125                     Call ChangerFournisseurRetour
130                   End If
135                 End If
140               End If
145             Else
150               Call ModifierTexte
155             End If
160           End If
165         End If
170       End If
175     End If

180     Exit Sub

AfficherErreur:

185     woups "frmProjSoumElec", "lvwSoumission_DblClick", Err, Erl
End Sub

Private Sub AjouterPrix()

5       On Error GoTo AfficherErreur

10      Call ViderChamps_frs

        'Rempli le combo des fournisseurs
15      Call RemplirComboFournisseur

20      cmbfrs.Locked = False

25      m_bMauvaisPrix = False

        'Positionne le frame
30      fraPrixPiece.Top = lvwSoumission.Top + 500
                
        'Montre le frame
35      fraPrixPiece.Visible = True

        'Met le numéro de la pièce dans le tag
40      fraPrixPiece.Tag = lvwSoumission.SelectedItem.Index
                  
        'Donne le focus au combo
45      Call cmbfrs.SetFocus

50      Exit Sub

AfficherErreur:

55      woups "frmProjSoumElec", "AjouterPrix", Err, Erl
End Sub

Private Sub ModifierTexte()

5       On Error GoTo AfficherErreur

10      Dim sTexte As String

15      sTexte = InputBox("Quel est le nouveau texte?", , lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_DESCR))

20      If sTexte <> "" Then
25        If Len(sTexte) > 255 Then
30          Call MsgBox("Le texte ne pas dépasser 255 caractères!", vbOKOnly, "Erreur")
35        Else
40          lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_DESCR) = sTexte
45        End If
50      End If

55      Exit Sub

AfficherErreur:

60      woups "frmProjSoumElec", "ModifierTexte", Err, Erl
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

145     woups "frmProjSoumElec", "ModifierSousSection", Err, Erl
End Sub

Private Sub ChangerFournisseurRetour()

5       On Error GoTo AfficherErreur

'10      m_bPieceInutile = True
15      m_bRecherchePiece = False
20      m_bChangementFRS = True

25      Call AfficherListeFournisseurs

30      If lvwfournisseur.ListItems.count = 0 Then
35        Call MsgBox("Il n'y a aucun fournisseur pour cette pièce!", vbOKOnly, "Erreur")
 
40        Exit Sub
45      Else
50        frafournisseur.Visible = True
55      End If

60      Exit Sub

AfficherErreur:

65      woups "frmProjSoumElec", "ChangerFournisseurRetour", Err, Erl
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

55      Exit Sub

AfficherErreur:

60      woups "frmProjSoumElec", "lvwSoumission_ItemCheck", Err, Erl
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
115                 If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PROFIT) = "" Then
120                   lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PROFIT) = " "
125                 End If

130                 If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PROFIT).Tag <> "EXTRA" Then
135                   bAfficherMenu = True
140                 Else
145                   bAfficherMenu = False
150                 End If
155               End If
160             Else
165               If lvwSoumission.DropHighlight.Selected = True Then
170                 If lvwSoumission.SelectedItem.Tag <> "" Then
175                   If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PROFIT).Tag <> "EXTRA" Then
180                     If g_bModificationFacturation = True Then
185                       bAfficherMenu = True
190                     Else
195                       bAfficherMenu = False
200                     End If
205                   Else
210                     bAfficherMenu = False
215                   End If
220                 Else
225                   bAfficherMenu = False
230                 End If
235               End If
240             End If
245           Else
250             bAfficherMenu = False
255           End If

260           If bAfficherMenu = True Then
265             Call RemplirOptionsMenuRightClick(iNbreSelected)

270             Call PopupMenu(mnuRightClick)
275           End If
280         End If
285       Else
290         If Shift <> vbCtrlMask And Shift <> vbShiftMask Then
295           Set lvwSoumission.DropHighlight = Nothing
300         End If
305       End If
310     End If

315     Exit Sub

AfficherErreur:

320     woups "frmProjSoumElec", "lvwSoumission_MouseDown", Err, Erl
End Sub

Private Sub RemplirOptionsMenuRightClick(ByVal iNbreSelected As Integer)

5       On Error GoTo AfficherErreur

10      Dim bFacturer        As Boolean
15      Dim bNC              As Boolean
20      Dim bDateRequise     As Boolean
25      Dim bCommentaire     As Boolean
30      Dim bID              As Boolean
35      Dim bMauvaisPrix     As Boolean
40      Dim bMaterielInutile As Boolean
45      Dim bTexte           As Boolean
50      Dim bSousSection     As Boolean
55      Dim bFournisseur     As Boolean
60      Dim bAnnulerCommande As Boolean
65      Dim bSupprimer       As Boolean
70      Dim bAjouterPrix     As Boolean
75      Dim bSortieMagasin   As Boolean
80      Dim bChangerQuantite As Boolean

85      If iNbreSelected > 1 Then
90        If m_eType = TYPE_PROJET Then
95          If g_bModificationFacturation = True Then
100           bFacturer = True
105           bNC = True
110           bSupprimer = True
115         End If
120       Else
125         bSupprimer = True
130       End If
135     Else
          'Si c'est une sous-section
140       If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) = "" Then
145         bSousSection = True
150       Else
155         If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) = "Texte" Or lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) = "Text" Then
160           bTexte = True

165           If (m_eType = TYPE_SOUMISSION) Or ((m_eType = TYPE_PROJET) And (Right$(txtNoProjSoum.Text, 2) > 19)) Then
170             bSupprimer = True
175           End If
180         Else
185           If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor = -2147483640 Then
190             lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor = 0
195           End If

200           Select Case lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor
                Case COLOR_ORANGE:
205               If g_bModificationFacturation = True Then
210                 bFacturer = True
215                 bNC = True
220               End If

225               bID = True
230               bDateRequise = True
235               bCommentaire = True
240               bAnnulerCommande = True
245               bMauvaisPrix = True
               
                Case COLOR_BRUN:
250               bCommentaire = True
255               bFournisseur = True

260               If (m_eType = TYPE_SOUMISSION) Or ((m_eType = TYPE_PROJET) And (Right$(txtNoProjSoum.Text, 2) > 19)) Then
265                 bSupprimer = True
270               End If

                Case COLOR_GRIS:
275               If g_bModificationFacturation = True Then
280                 bFacturer = True
285                 bNC = True
290               End If

295               bCommentaire = True
300               bID = True
305               bMauvaisPrix = True
310               bMaterielInutile = True

                Case COLOR_VERT_FORET:
315               bCommentaire = True

320               If (m_eType = TYPE_SOUMISSION) Or ((m_eType = TYPE_PROJET) And (Right$(txtNoProjSoum.Text, 2) > 19)) Then
325                 bSupprimer = True
330               End If

                Case COLOR_ROUGE:
335               bCommentaire = True

                Case COLOR_MAGENTA:
340               bCommentaire = True
345               bAjouterPrix = True

350               If m_eType = TYPE_PROJET Then
355                 bID = True
360               End If

365               If m_eType = TYPE_SOUMISSION Then
370                 bChangerQuantite = True
375               End If

                Case COLOR_NOIR:
380               If m_eType = TYPE_PROJET Then
385                 If g_bModificationFacturation = True Then
390                   bFacturer = True
395                   bNC = True
400                 End If

405                 bID = True
410                 bMaterielInutile = True
415                 bSortieMagasin = True
420               Else
425                 bChangerQuantite = True
430               End If

435               bCommentaire = True
440               bMauvaisPrix = True
445               bFournisseur = True
450               bSupprimer = True
455           End Select
460         End If
465       End If
470     End If

        'Pour empeche que tous les éléments deviennent invisible, je les mets visible au
        'début
475     mnuFacturer.Visible = True
480     mnuNC.Visible = True
485     mnuDateRequise.Visible = True
490     mnuCommentaire.Visible = True
495     mnuID.Visible = True
500     mnuMauvaisPrix.Visible = True
505     mnuInutile.Visible = True
510     mnuTexte.Visible = True
515     mnuChangerSS.Visible = True
520     mnuFournisseur.Visible = True
525     mnuAnnulerCommande.Visible = True
530     mnuSupprimer.Visible = True
535     mnuAjouterPrix.Visible = True
540     mnuSortieMagasin.Visible = True
545     mnuQuantite.Visible = True

550     mnuFacturer.Visible = bFacturer
555     mnuNC.Visible = bNC
560     mnuDateRequise.Visible = bDateRequise
565     mnuCommentaire.Visible = bCommentaire
570     mnuID.Visible = bID
575     mnuMauvaisPrix.Visible = bMauvaisPrix
580     mnuInutile.Visible = bMaterielInutile
585     mnuTexte.Visible = bTexte
590     mnuChangerSS.Visible = bSousSection
595     mnuFournisseur.Visible = bFournisseur
600     mnuAnnulerCommande.Visible = bAnnulerCommande
605     mnuSupprimer.Visible = bSupprimer
610     mnuAjouterPrix.Visible = bAjouterPrix
615     mnuSortieMagasin.Visible = bSortieMagasin
620     mnuQuantite.Visible = bChangerQuantite

625     Exit Sub

AfficherErreur:

630     woups "frmProjSoumElec", "RemplirOptionsMenuRightClick", Err, Erl
End Sub

Private Sub mnuAjouterPrix_Click()

5       On Error GoTo AfficherErreur

10      Call AjouterPrix

15      Call EnleverSelection

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumElec", "mnuAjouterPrix_Click", Err, Erl
End Sub

Private Sub mnuAnnulerCommande_Click()

5       On Error GoTo AfficherErreur

10      Call AnnulerCommande

15      Call EnleverSelection

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumElec", "mnuAnnulerCommande_Click", Err, Erl
End Sub

Private Sub mnuChangerSS_Click()

5       On Error GoTo AfficherErreur

10      Call ModifierSousSection

15      Call EnleverSelection

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumElec", "mnuChangerSS_Click", Err, Erl
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

70      woups "frmProjSoumElec", "mnuDateRequise_Click", Err, Erl
End Sub

Private Sub mnuCommentaire_Click()

5       On Error GoTo AfficherErreur

10      txtcommentaire.Text = lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_COMMENTAIRE)

15      fraCommentaire.Top = lvwSoumission.Top

20      fraCommentaire.Visible = True

25      Exit Sub

AfficherErreur:

30      woups "frmProjSoumElec", "mnuCommentaire_Click", Err, Erl
End Sub

Private Sub mnuFacturer_Click()

5       On Error GoTo AfficherErreur

10      Call FacturerDate

15      Call EnleverSelection

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumElec", "mnuFacturer_Click", Err, Erl
End Sub

Private Sub mnuFournisseur_Click()

5       On Error GoTo AfficherErreur

10      Call ChangerFournisseurRetour

15      Call EnleverSelection

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumElec", "mnuFournisseur_Click", Err, Erl
End Sub

Private Sub mnuID_Click()

5       On Error GoTo AfficherErreur

10      lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_ID) = InputBox("Quel est l'ID", , lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_ID))

15      Call EnleverSelection

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumElec", "mnuID_Click", Err, Erl
End Sub

Private Sub mnuInutile_Click()

5       On Error GoTo AfficherErreur

10      Call MaterielInutile

15      Call EnleverSelection

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumElec", "mnuInutile_Click", Err, Erl
End Sub

Private Sub mnuMauvaisPrix_Click()

5       On Error GoTo AfficherErreur

10      Call MauvaisPrix

15      Call EnleverSelection

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumElec", "mnuMauvaisPrix", Err, Erl
End Sub

Private Sub mnuNC_Click()

5       On Error GoTo AfficherErreur

10      Call FacturerNC

15      Call EnleverSelection

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumElec", "mnuNC_Click", Err, Erl
End Sub

Private Sub mnuQuantite_Click()

5       On Error GoTo AfficherErreur

10      Call ChangerQuantite

15      Call EnleverSelection

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumElec", "mnuQuantite_Click", Err, Erl
End Sub

Private Sub mnuSortieMagasin_Click()

5       On Error GoTo AfficherErreur

10      Call SortieMagasin

15      Call EnleverSelection

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumElec", "mnuSortieMagasin_Click", Err, Erl
End Sub

Private Sub mnuSupprimer_Click()
  
5       On Error GoTo AfficherErreur

10      Call EffacerItemListViewSoumission

15      Call EnleverSelection

20      Exit Sub
  
AfficherErreur:

25      woups "frmProjSoumElec", "mnuSupprimer_Click", Err, Erl
End Sub

Private Sub mnuTexte_Click()

5       On Error GoTo AfficherErreur

10      Call ModifierTexte

15      Call EnleverSelection

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumElec", "mnuTexte_Click", Err, Erl
End Sub

Private Sub EnleverSelection()

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer

15      Set lvwSoumission.DropHighlight = Nothing

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumElec", "EnleverSelection", Err, Erl
End Sub

Private Sub mvwDate_LostFocus()

5       On Error GoTo AfficherErreur
        
        'Quand le calendrier perd le focus, il faut l'enlever
10      mvwDate.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElec", "mvwDate_LostFocus", Err, Erl
End Sub

Private Sub mvwDateFacturation_LostFocus()

5       On Error GoTo AfficherErreur

        'Quand le calendrier perd le focus, il faut l'enlever
10      mvwDateFacturation.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElec", "mvwDateFacturation_LostFocus", Err, Erl
End Sub

Private Sub mvwDate_DateClick(ByVal DateClicked As Date)

5       On Error GoTo AfficherErreur

        'Affiche la date dans le TextBox sous le format AAAA-MM-JJ
10      txtDelais.Text = ConvertDate(DateClicked)
  
        'Enlever le calendrier
15      mvwDate.Visible = False

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumElec", "mvwDate_DateClick", Err, Erl
End Sub

Private Sub mvwDateFacturation_DateClick(ByVal DateClicked As Date)

5       On Error GoTo AfficherErreur

        'Affiche la date dans le TextBox sous le format AAAA-MM-JJ
10      txtDateFacturation.Text = ConvertDate(DateClicked)

        'Enlever le calendrier
15      mvwDateFacturation.Visible = False

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumElec", "mvwDateFacturation_DateClick", Err, Erl
End Sub

Private Sub cmdEnregistrer_Click()

5       On Error GoTo AfficherErreur

10      Dim objControl As Control
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
                 objControl.Name <> "txtDelais" And _
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
             
        'Vérification du transport
155     If cmbtransport.ListIndex = -1 Then
160       Call MsgBox("Vous devez choisir le transport!", vbOKOnly, "Erreur")
    
165       Exit Sub
170     End If

175     If m_eType = TYPE_SOUMISSION Then
180       If m_sTempsTest = 0 Or m_sTempsDessin = 0 Then
185         If MsgBox("Les temps de dessin ou de test sont vides" & vbNewLine & "Voulez - vous l'enregistrer?", vbYesNo) = vbNo Then
190           Exit Sub
195         End If
200       End If
205     End If

210     Screen.MousePointer = vbHourglass

215     If BackupPieces(txtNoProjSoum.Text) = False Then
220       If m_eType = TYPE_PROJET Then
225         sMessage = "Une erreur est survenue lors de la copie de sauvegarde du projet en cours!"
230       Else
235         sMessage = "Une erreur est survenue lors de la copie de sauvegarde de la soumission en cours!"
240       End If

245       sMessage = sMessage & vbNewLine & vbNewLine & "Voulez-vous continuer ?"

250       Screen.MousePointer = vbDefault

255       If MsgBox(sMessage, vbYesNo) = vbNo Then
260         Exit Sub
265       Else
270         Screen.MousePointer = vbHourglass
275       End If
280     End If
        
        'Enregistre la soumission
285     Call EnregistrerProjSoum(txtNoProjSoum.Text)
  
290     Call OuvrirProjSoum(False)
   
        'Remet en mode inactif
295     Call AfficherControles(MODE_INACTIF)

300     m_bEnregistrement = True
  
        'Affiche la soumission actuel
305     Call AfficherProjSoum(txtNoProjSoum.Text)

310     m_bEnregistrement = False
  
315     Screen.MousePointer = vbDefault

320     Exit Sub

AfficherErreur:

325     woups "frmProjSoumElec", "cmdEnregistrer_Click", Err, Erl
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
40      Dim sTempsFab     As String
45      Dim sTotalPiece   As String
50      Dim sImprevue     As String
55      Dim sTotalTemps   As String
60      Dim sManuel       As String
65      Dim iCompteur     As Integer
70      Dim iIndexFacture As Integer
75      Dim collFacture   As Collection
80      Dim bExiste       As Boolean

85      Set collFacture = New Collection

90      Call g_connData.Execute("DELETE * FROM GRB_Projet_Modif WHERE IDProjet = '" & sNoProjet & "' AND Type = 'E' AND TypeModif = 'FACTURATION'")

95      If lvwSoumission.ListItems.count > 0 Then
100       Set rstModif = New ADODB.Recordset
105       Set rstEmploye = New ADODB.Recordset

110       For iCompteur = 1 To lvwSoumission.ListItems.count
115         bExiste = False

120         sNoFacture = lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_FACTURATION)

125         If Trim$(sNoFacture) <> "" Then
130           For iIndexFacture = 1 To collFacture.count
135             If collFacture(iIndexFacture) = sNoFacture Then
140               bExiste = True

145               Exit For
150             End If
155           Next

160           If bExiste = False Then
165             Call collFacture.Add(sNoFacture)

170             Call CalculerPrixFacturation(sNoFacture, sCommission, sPrixTotal, sProfit, sTempsFab, sTotalPiece, sImprevue, sTotalTemps, sManuel)

175             Call rstModif.Open("SELECT * FROM GRB_Projet_Modif WHERE [Date] = '" & Replace(sNoFacture, "F-", "") & "' AND TypeModif = 'FACTURATION'", g_connData, adOpenDynamic, adLockOptimistic)

180             If rstModif.EOF Then
185               Call rstModif.AddNew
190             End If

195             rstModif.Fields("IDProjet") = txtNoProjSoum.Text

200             Call rstEmploye.Open("SELECT * FROM GRB_Employés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)

205             rstModif.Fields("NoEmployé") = rstEmploye.Fields("NoEmploye")

210             Call rstEmploye.Close

215             rstModif.Fields("Date") = Replace(sNoFacture, "F-", "")
220             rstModif.Fields("Heure") = " "
225             rstModif.Fields("Type") = "E"
230             rstModif.Fields("TypeModif") = "FACTURATION"
235             rstModif.Fields("Valeur") = sPrixTotal

240             Call rstModif.Update

245             Call rstModif.Close
250           End If
255         End If
260       Next

265       Set rstModif = Nothing
270       Set rstEmploye = Nothing
275     End If

280     Exit Sub

AfficherErreur:

285     woups "frmProjSoumElec", "EnregistrerFACT", Err, Erl
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
190       rstProjSoumBackup.Fields("sousSection") = rstProjSoum.Fields("sousSection")
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

385     woups "frmProjSoumElec", "BackupPieces", Err, Erl
End Function

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
         
145     Set rstProjSoum = New ADODB.Recordset
150     Set rstPiece = New ADODB.Recordset
155     Set rstEmploye = New ADODB.Recordset
160     Set rstModif = New ADODB.Recordset
165     Set rstOuvert = New ADODB.Recordset
170     Set rstSection = New ADODB.Recordset

175     Set collExtra = New Collection

180     Call rstEmploye.Open("SELECT noEmploye FROM GRB_Employés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)

        'Si c'est un projet
185     If m_eType = TYPE_PROJET Then
190       sTable = "GRB_ProjetElec"
195       sTableModif = "GRB_Projet_Modif"
200       sTablePiece = "GRB_Projet_Pieces"
205       sChamps = "IDProjet"
210     Else
215       sTable = "GRB_SoumissionElec"
220       sTableModif = "GRB_Soumission_Modif"
225       sTablePiece = "GRB_Soumission_Pieces"
230       sChamps = "IDSoumission"
235     End If

        'Si c'est un ajout
240     If m_bModeAjout = True Then
          'On ouvre le recordset selon le type
245       Call rstProjSoum.Open("SELECT * FROM " & sTable & " WHERE " & sChamps & " = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)

250       If m_eType = TYPE_PROJET Then
255         If rstProjSoum.EOF Then
260           bAjoutCommande = True
265         Else
270           bAjoutCommande = False
275         End If
280       Else
285         bAjoutCommande = False
290       End If

295       Call rstProjSoum.AddNew

300       If m_eType = TYPE_PROJET Then
305         rstProjSoum.Fields("LiaisonChargeable") = m_sLiaison
310       End If

315       rstProjSoum.Fields("Creer") = ConvertDate(Date)
320       rstProjSoum.Fields("Creer_Par") = rstEmploye.Fields("noEmploye")

325       rstProjSoum.Fields(sChamps) = sNoProjSoum

330       If m_eType = TYPE_PROJET Then
335         rstProjSoum.Fields("IDSoumission") = txtNoSoumission.Text
340       End If

345       Call rstOuvert.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)

350       If rstOuvert.EOF Then
355         Call rstOuvert.AddNew

360         rstOuvert.Fields("IDProjSoum") = sNoProjSoum
365         rstOuvert.Fields("NoClient") = cmbclient.ItemData(cmbclient.ListIndex)
370         rstOuvert.Fields("Description") = txtProjet.Text
375         rstOuvert.Fields("DateOuverture") = ConvertDate(Date)
380         rstOuvert.Fields("Ouvert") = True
    
385         If m_eType = TYPE_PROJET Then
390           rstOuvert.Fields("Type") = "P"
395         Else
400           rstOuvert.Fields("Type") = "S"
405         End If

410         Call rstOuvert.Update
415       End If
    
420       Call rstOuvert.Close
425       Set rstOuvert = Nothing

430       m_bModeAjout = False
435     Else
440       Call EnregistrerSuppression

445       Call rstProjSoum.Open("SELECT * FROM " & sTable & " WHERE " & sChamps & " = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)
      
450       Call rstModif.Open("SELECT * FROM " & sTableModif, g_connData, adOpenDynamic, adLockOptimistic)
      
455       Call rstModif.AddNew
      
460       rstModif.Fields("Type") = "E"
465       rstModif.Fields(sChamps) = sNoProjSoum
470       rstModif.Fields("noEmployé") = rstEmploye.Fields("noEmploye")
475       rstModif.Fields("Date") = ConvertDate(Date)
480       rstModif.Fields("Heure") = Time
485       rstModif.Fields("TypeModif") = "MODIFICATION"
    
490       Call rstModif.Update
    
495       Call rstModif.Close
500       Set rstModif = Nothing

505       Call rstOuvert.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)
    
510       If rstOuvert.Fields("NoClient") <> cmbclient.ItemData(cmbclient.ListIndex) Then
515         rstOuvert.Fields("NoClient") = cmbclient.ItemData(cmbclient.ListIndex)

520         Set rstPunch = New ADODB.Recordset

525         Call rstPunch.Open("SELECT * FROM GRB_Punch WHERE NoProjet = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)

530         If Not rstPunch.EOF Then
535           If MsgBox("Le client a été modifié, voulez-vous changer les punch de ce projet ?", vbYesNo) = vbYes Then

540             Do While Not rstPunch.EOF
545               rstPunch.Fields("NoClient") = cmbclient.ItemData(cmbclient.ListIndex)

550               Call rstPunch.Update

555               Call rstPunch.MoveNext
560             Loop
565           End If
570         End If

575         Call rstPunch.Close
580         Set rstPunch = Nothing
585       End If

590       rstOuvert.Fields("Description") = txtProjet.Text

595       Call rstOuvert.Update

600       Call rstOuvert.Close
605       Set rstOuvert = Nothing

          'Si c'est une modification, il faut effacer les pieces et remplir les nouvelles
610       Call g_connData.Execute("DELETE * FROM " & sTablePiece & " WHERE " & sChamps & " = '" & sNoProjSoum & "' AND Type = 'E'")

615       If m_eType = TYPE_PROJET Then
620         If Right$(sNoProjSoum, 2) >= 60 And Right$(sNoProjSoum, 2) <= 98 Then
625           Call g_connData.Execute("DELETE * FROM GRB_Projet_Pieces WHERE IDProjet = '" & Left$(sNoProjSoum, 6) & "-" & rstProjSoum.Fields("LiaisonChargeable") & "' AND Type = 'E' AND (PieceExtraChargeable = True OR PieceExtraNonChargeable = True) AND Provenance = '" & Right$(txtNoProjSoum.Text, 2) & "'")
630         End If
635       End If
640     End If
              
        'Enregistrement de la soumission
        'Pour savoir que c'est une soumission ou un projet électrique
645     rstProjSoum.Fields("IDClient") = cmbclient.ItemData(cmbclient.ListIndex)
650     rstProjSoum.Fields("IDContact") = cmbContact.ItemData(cmbContact.ListIndex)
655     rstProjSoum.Fields("description") = txtProjet.Text
660     rstProjSoum.Fields("NbreManuel") = txtNbreManuel.Text
665     rstProjSoum.Fields("transport") = cmbtransport.Text
670     rstProjSoum.Fields("CSA") = chkCSA.Value
675     rstProjSoum.Fields("CUL") = chkCUL.Value
680     rstProjSoum.Fields("UL") = chkUL.Value
685     rstProjSoum.Fields("CUR") = chkCUR.Value
690     rstProjSoum.Fields("UR") = chkUR.Value
695     rstProjSoum.Fields("CE") = chkCE.Value

700     If txtDelais.Text <> "" Then
705       rstProjSoum.Fields("Delais") = txtDelais.Text
710     Else
715       rstProjSoum.Fields("Delais") = " "
720     End If

725     If m_eType = TYPE_SOUMISSION Then
730       rstProjSoum.Fields("TempsDessin") = m_sTempsDessin
735       rstProjSoum.Fields("TempsFabrication") = m_sTempsFabrication
740       rstProjSoum.Fields("TempsAssemblage") = m_sTempsAssemblage
745       rstProjSoum.Fields("TempsProgInterface") = m_sTempsProgInterface
750       rstProjSoum.Fields("TempsProgAutomate") = m_sTempsProgAutomate
755       rstProjSoum.Fields("TempsProgRobot") = m_sTempsProgRobot
760       rstProjSoum.Fields("TempsVision") = m_sTempsVision
765       rstProjSoum.Fields("TempsTest") = m_sTempsTest
770       rstProjSoum.Fields("TempsInstallation") = m_sTempsInstallation
775       rstProjSoum.Fields("TempsMiseService") = m_sTempsMiseService
780       rstProjSoum.Fields("TempsFormation") = m_sTempsFormation
785       rstProjSoum.Fields("TempsGestion") = m_sTempsGestion
790       rstProjSoum.Fields("TempsShipping") = m_sTempsShipping
795     End If

800     rstProjSoum.Fields("NbrePersonne") = m_sNbrePersonne
805     rstProjSoum.Fields("TempsHebergement") = m_sTempsHebergement
810     rstProjSoum.Fields("TempsRepas") = m_sTempsRepas
815     rstProjSoum.Fields("TempsTransport") = m_sTempsTransport
820     rstProjSoum.Fields("TempsUniteMobile") = m_sTempsUniteMobile
825     rstProjSoum.Fields("PrixEmballage") = m_sPrixEmballage

830     rstProjSoum.Fields("TauxHebergement1") = m_sTauxHebergement1
835     rstProjSoum.Fields("TauxHebergement2") = m_sTauxHebergement2
840     rstProjSoum.Fields("TauxRepas") = m_sTauxRepas
845     rstProjSoum.Fields("TauxTransport") = m_sTauxTransport
850     rstProjSoum.Fields("TauxUniteMobile") = m_sTauxUniteMobile

855     rstProjSoum.Fields("TauxDessin") = m_sTauxDessin
860     rstProjSoum.Fields("TauxFabrication") = m_sTauxFabrication
865     rstProjSoum.Fields("TauxAssemblage") = m_sTauxAssemblage
870     rstProjSoum.Fields("TauxProgInterface") = m_sTauxProgInterface
875     rstProjSoum.Fields("TauxProgAutomate") = m_sTauxProgAutomate
880     rstProjSoum.Fields("TauxProgRobot") = m_sTauxProgRobot
885     rstProjSoum.Fields("TauxVision") = m_sTauxVision
890     rstProjSoum.Fields("TauxTest") = m_sTauxTest
895     rstProjSoum.Fields("TauxInstallation") = m_sTauxInstallation
900     rstProjSoum.Fields("TauxMiseService") = m_sTauxMiseService
905     rstProjSoum.Fields("TauxFormation") = m_sTauxFormation
910     rstProjSoum.Fields("TauxGestion") = m_sTauxGestion
915     rstProjSoum.Fields("TauxShipping") = m_sTauxShipping

920     rstProjSoum.Fields("imprevue") = m_sImprevue
925     rstProjSoum.Fields("commission") = m_sCommission
930     rstProjSoum.Fields("Profit") = m_sProfit
935     rstProjSoum.Fields("SansTemps") = m_bSansTemps
940     rstProjSoum.Fields("CheminPhotos") = txtCheminPhotos.Text
945     rstProjSoum.Fields("MontantForfait") = txtForfait.Text
950     rstProjSoum.Fields("InitialeForfait") = Trim$(Replace(lblForfaitInitiale.Caption, "Par :", ""))

955     If Not IsNull(rstProjSoum.Fields("NbrePersonne")) Then
960       dblNbrePers = CDbl(rstProjSoum.Fields("NbrePersonne"))
965     Else
970       dblNbrePers = 0
975     End If

980     If Not IsNull(rstProjSoum.Fields("TempsHebergement")) Then
985       dblJoursHebergement = CDbl(rstProjSoum.Fields("TempsHebergement"))
990     Else
995       dblJoursHebergement = 0
1000    End If

1005    If Not IsNull(rstProjSoum.Fields("TempsRepas")) Then
1010      dblJoursRepas = CDbl(rstProjSoum.Fields("TempsRepas"))
1015    Else
1020      dblJoursRepas = 0
1025    End If

1030    If Not IsNull(rstProjSoum.Fields("TauxHebergement1")) Then
1035      dblHebergement1 = CDbl(rstProjSoum.Fields("TauxHebergement1"))
1040    Else
1045      dblHebergement1 = 0
1050    End If

1055    If Not IsNull(rstProjSoum.Fields("TauxHebergement2")) Then
1060      dblHebergement2 = CDbl(rstProjSoum.Fields("TauxHebergement2"))
1065    Else
1070      dblHebergement2 = 0
1075    End If

1080    If Not IsNull(rstProjSoum.Fields("TauxRepas")) Then
1085      dblRepas = CDbl(rstProjSoum.Fields("TauxRepas"))
1090    Else
1095      dblRepas = 0
1100    End If

1105    rstProjSoum.Fields("TotalRepas") = dblNbrePers * dblJoursRepas * dblRepas

1110    dblTotalHebergement = 0

1115    Do While dblNbrePers > 0
1120      If dblNbrePers >= 2 Then
1125        dblTotalHebergement = dblTotalHebergement + (dblJoursHebergement * dblHebergement2)

1130        dblNbrePers = dblNbrePers - 2
1135      Else
1140        dblTotalHebergement = dblTotalHebergement + (dblJoursHebergement * dblHebergement1)

1145        dblNbrePers = dblNbrePers - 1
1150      End If
1155    Loop

1160    rstProjSoum.Fields("TotalHebergement") = dblTotalHebergement

1165    If bAjoutCommande = True Then
1170      rstProjSoum.Fields("ProchaineCommande") = 1
1175    End If

        'Si c'est un projet, il faut enregistrer le prix de réception
1180    If m_eType = TYPE_PROJET Then
1185      rstProjSoum.Fields("PrixRéception") = txtPrixReception.Text
1190    End If

1195    If IsNumeric(txtPrixManuel.Text) Then
1200      rstProjSoum.Fields("Total_Manuel") = txtPrixManuel.Text
1205    Else
1210      rstProjSoum.Fields("Total_Manuel") = "0"
1215    End If

1220    rstProjSoum.Fields("total_Commission") = txtCommission.Text
1225    rstProjSoum.Fields("Total_Profit") = txtProfit.Text
1230    rstProjSoum.Fields("PrixTotal") = txtPrixTotal.Text
1235    rstProjSoum.Fields("Total_piece") = txtTotalPieces.Text
1240    rstProjSoum.Fields("total_imprevue") = txtImprevus.Text
   
1245    rstProjSoum.Fields("PrixTotal") = txtPrixTotal.Text
 
1250    rstPiece.CursorLocation = adUseServer

1255    If m_eType = TYPE_PROJET Then
1260      Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces", g_connData, adOpenDynamic, adLockOptimistic)
1265    Else
1270      Call rstPiece.Open("SELECT * FROM GRB_Soumission_Pieces", g_connData, adOpenDynamic, adLockOptimistic)
1275    End If

1280    If m_eType = TYPE_PROJET Then
1285      If g_bModificationFacturation = True Then
1290        Call EnregistrerFACT(sNoProjSoum)
1295      End If
1300    End If

        'Enregistrement des pièces
1305    For iCompteur = 1 To lvwSoumission.ListItems.count
1310      If lvwSoumission.ListItems(iCompteur).Tag <> vbNullString Then
1315        If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> vbNullString Then
1320          Set itmPiece = lvwSoumission.ListItems(iCompteur)
     
1325          Call rstPiece.AddNew

1330          If m_eType = TYPE_PROJET Then
1335            rstPiece.Fields("IDProjet") = sNoProjSoum
1340          Else
1345            rstPiece.Fields("IDSoumission") = sNoProjSoum
1350          End If
       
1355          rstPiece.Fields("Type") = "E"
       
1360          If itmPiece.Checked = True Then
1365            rstPiece.Fields("Visible") = True
1370          Else
1375            rstPiece.Fields("Visible") = False
1380          End If

1385          If m_eType = TYPE_PROJET Then
1390            rstPiece.Fields("Facturation") = itmPiece.SubItems(I_COL_SOUM_FACTURATION)

1395            If itmPiece.SubItems(I_COL_SOUM_FACTURATION) = "" Then
1400              itmPiece.SubItems(I_COL_SOUM_FACTURATION) = " "
1405            End If

1410            If itmPiece.SubItems(I_COL_SOUM_DATE_COMMANDE) = "" Then
1415              itmPiece.SubItems(I_COL_SOUM_DATE_COMMANDE) = " "
1420            End If

1425            rstPiece.Fields("NoRetour") = itmPiece.ListSubItems(I_COL_SOUM_DATE_COMMANDE).Tag

1430            rstPiece.Fields("DateRéception") = itmPiece.ListSubItems(I_COL_SOUM_PRIX_NET).Tag
1435          End If
               
1440          rstPiece.Fields("IDSection") = itmPiece.Tag
1445          rstPiece.Fields("NumItem") = Trim$(itmPiece.SubItems(I_COL_SOUM_PIECE))
1450          rstPiece.Fields("Qté") = Replace(itmPiece.Text, "*", vbNullString)

1455          If itmPiece.SubItems(I_COL_SOUM_PIECE) = "Texte" Or itmPiece.SubItems(I_COL_SOUM_PIECE) = "Text" Then
1460            rstPiece.Fields("DESC_EN") = ""
1465            rstPiece.Fields("DESC_FR") = Trim$(itmPiece.SubItems(I_COL_SOUM_DESCR))
1470          Else
1475            If m_eLangage = ANGLAIS Then
1480              rstPiece.Fields("DESC_EN") = Trim$(itmPiece.SubItems(I_COL_SOUM_DESCR))
1485              rstPiece.Fields("DESC_FR") = Trim$(itmPiece.ListSubItems(I_COL_SOUM_DESCR).Tag)
1490            Else
1495              rstPiece.Fields("DESC_FR") = Trim$(itmPiece.SubItems(I_COL_SOUM_DESCR))
1500              rstPiece.Fields("DESC_EN") = Trim$(itmPiece.ListSubItems(I_COL_SOUM_DESCR).Tag)
1505            End If
1510          End If

1515          rstPiece.Fields("Manufact") = Trim$(itmPiece.SubItems(I_COL_SOUM_MANUFACT))
1520          rstPiece.Fields("Prix_list") = Conversion(itmPiece.SubItems(I_COL_SOUM_PRIX_LIST), MODE_PAS_FORMAT, 4)

1525          If Trim$(itmPiece.SubItems(I_COL_SOUM_ESCOMPTE)) <> "" Then
1530            rstPiece.Fields("Escompte") = Conversion(Replace(itmPiece.SubItems(I_COL_SOUM_ESCOMPTE), "%", "") / 100, MODE_PAS_FORMAT)
1535          Else
1540            rstPiece.Fields("Escompte") = ""
1545          End If

1550          rstPiece.Fields("Prix_net") = Conversion(itmPiece.SubItems(I_COL_SOUM_PRIX_NET), MODE_PAS_FORMAT, 4)

1555          rstPiece.Fields("OrdreSection") = itmPiece.ListSubItems(I_COL_SOUM_MANUFACT).Tag
1560          rstPiece.Fields("NuméroLigne") = iCompteur
            
              'Si le listItem est COLOR_ORANGE
1565          If itmPiece.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_ORANGE Then
1570            rstPiece.Fields("Commandé") = True
1575          Else
1580            rstPiece.Fields("Commandé") = False
1585          End If

1590          If itmPiece.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_GRIS Then
1595            rstPiece.Fields("Recu") = True
1600          Else
1605            rstPiece.Fields("Recu") = False
1610          End If

1615          If itmPiece.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_ROUGE Then
1620            rstPiece.Fields("Retour") = True
1625          Else
1630            rstPiece.Fields("Retour") = False
1635          End If

1640          If itmPiece.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_VERT_FORET And itmPiece.ListSubItems(I_COL_SOUM_PIECE).Bold = True Then
1645            rstPiece.Fields("CommandeAnnulée") = True
1650          Else
1655            rstPiece.Fields("CommandeAnnulée") = False
1660          End If

1665          If itmPiece.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_BRUN Then
1670            rstPiece.Fields("MatérielInutile") = True
1675          Else
1680            rstPiece.Fields("MatérielInutile") = False
1685          End If
   
1690          If itmPiece.SubItems(I_COL_SOUM_DISTRIB) <> vbNullString Then
1695            rstPiece.Fields("IDFRS") = itmPiece.ListSubItems(I_COL_SOUM_DISTRIB).Tag
1700          End If
     
1705          rstPiece.Fields("Temps") = Trim$(itmPiece.SubItems(I_COL_SOUM_TEMPS))
1710          rstPiece.Fields("Temps_Total") = itmPiece.SubItems(I_COL_SOUM_MONTAGE)
1715          rstPiece.Fields("Prix_Total") = Conversion(itmPiece.SubItems(I_COL_SOUM_TOTAL), MODE_PAS_FORMAT)
1720          rstPiece.Fields("Profit_argent") = Conversion(itmPiece.SubItems(I_COL_SOUM_PROFIT), MODE_PAS_FORMAT)

1725          If Len(itmPiece.ListSubItems(I_COL_SOUM_PIECE).Tag) <= 50 Then
1730            rstPiece.Fields("SousSection") = itmPiece.ListSubItems(I_COL_SOUM_PIECE).Tag
1735          Else
1740            rstPiece.Fields("SousSection") = Left$(itmPiece.ListSubItems(I_COL_SOUM_PIECE).Tag, 50)
1745          End If
            
1750          If itmPiece.SubItems(I_COL_SOUM_PRIX_LIST) <> vbNullString Then
1755            If itmPiece.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag <> vbNullString Then
1760              rstPiece.Fields("PrixOrigine") = Replace(Round(CDbl(Replace(itmPiece.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag, ".", ",")), 4), ".", ",")
1765            Else
1770              rstPiece.Fields("PrixOrigine") = "0"
1775            End If
1780          Else
1785            rstPiece.Fields("PrixOrigine") = "0"
1790          End If

1795          If itmPiece.SubItems(I_COL_SOUM_TOTAL) <> "" Then
1800            rstPiece.Fields("Devise") = itmPiece.ListSubItems(I_COL_SOUM_TOTAL).Tag
1805          Else
1810            rstPiece.Fields("Devise") = ""
1815          End If
  
1820          If InStr(1, itmPiece.Text, "*") > 0 Then
1825            rstPiece.Fields("Quoté") = True
1830          Else
1835            rstPiece.Fields("Quoté") = False
1840          End If

1845          If m_eType = TYPE_PROJET Then
1850            If Trim$(itmPiece.SubItems(I_COL_SOUM_ID) <> vbNullString) Then
1855              rstPiece.Fields("ID") = itmPiece.SubItems(I_COL_SOUM_ID)
1860            End If

1865            rstPiece.Fields("DateCommande") = itmPiece.SubItems(I_COL_SOUM_DATE_COMMANDE)
1870            rstPiece.Fields("DateRequise") = itmPiece.SubItems(I_COL_SOUM_DATE_REQUISE)

1875            If itmPiece.SubItems(I_COL_SOUM_DATE_REQUISE) = "" Then
1880              itmPiece.SubItems(I_COL_SOUM_DATE_REQUISE) = " "
1885            End If

1890            rstPiece.Fields("DateRetour") = itmPiece.ListSubItems(I_COL_SOUM_DATE_REQUISE).Tag

1895            rstPiece.Fields("NomCommande") = itmPiece.SubItems(I_COL_SOUM_NOM_COMMANDE)

1900            rstPiece.Fields("NoSéquentiel") = itmPiece.SubItems(I_COL_SOUM_NO_SEQUENTIEL)
1905          End If

1910          If m_eType = TYPE_PROJET Then
1915            If itmPiece.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_ROSE Then
1920              rstPiece.Fields("PieceExtraNonChargeable") = True
1925              rstPiece.Fields("PieceExtraChargeable") = False
1930            Else
1935              If itmPiece.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_BLEU Then
1940                rstPiece.Fields("PieceExtraChargeable") = True
1945                rstPiece.Fields("PieceExtraNonChargeable") = False
1950              Else
1955                If Left$(itmPiece.ListSubItems(I_COL_SOUM_PROFIT).Tag, 6) = "RETOUR" Or Left$(itmPiece.ListSubItems(I_COL_SOUM_PROFIT).Tag, 10) = "ANNULATION" Then
1960                  sExtra = Right$(itmPiece.ListSubItems(I_COL_SOUM_PROFIT).Tag, 2)

1965                  If sExtra >= "80" And sExtra <= "98" Then
1970                    rstPiece.Fields("PieceExtraNonChargeable") = True
1975                    rstPiece.Fields("PieceExtraChargeable") = False
1980                  Else
1985                    rstPiece.Fields("PieceExtraChargeable") = True
1990                    rstPiece.Fields("PieceExtraNonChargeable") = False
1995                  End If
2000                End If
2005              End If
2010            End If

2015            If itmPiece.SubItems(I_COL_SOUM_PROVENANCE) <> "" Then
2020              rstPiece.Fields("Provenance") = Right$(itmPiece.SubItems(I_COL_SOUM_PROVENANCE), 2)
2025            Else
2030              If Left$(itmPiece.ListSubItems(I_COL_SOUM_PROFIT).Tag, 6) = "RETOUR" Or Left$(itmPiece.ListSubItems(I_COL_SOUM_PROFIT).Tag, 10) = "ANNULATION" Then
2035                rstPiece.Fields("Provenance") = sExtra
2040              End If
2045            End If
2050          End If

2055          rstPiece.Fields("Commentaire") = itmPiece.SubItems(I_COL_SOUM_COMMENTAIRE)

2060          Call rstPiece.Update

2065          If m_eType = TYPE_PROJET Then
2070            If Right$(txtNoProjSoum.Text, 2) <= 98 And Right$(txtNoProjSoum.Text, 2) >= 80 Then
2075              Call AjouterPiecesExtraDansJob(itmPiece, rstProjSoum.Fields("LiaisonChargeable"))
2080            Else
2085              If Right$(txtNoProjSoum.Text, 2) <= 79 And Right$(txtNoProjSoum.Text, 2) >= 60 Then
2090                Call AjouterPiecesExtraChargeableDansJob(itmPiece, rstProjSoum.Fields("LiaisonChargeable"))
2095              Else
2100                If Left$(itmPiece.ListSubItems(I_COL_SOUM_PROFIT).Tag, 6) = "RETOUR" Then
2105                  Call AjouterInutileDansExtra(itmPiece, sExtra)

2110                  bCalculExtra = True

2115                  bExiste = False

2120                  For iCompteurExtra = 1 To collExtra.count
2125                    If collExtra(iCompteurExtra) = Right$(itmPiece.ListSubItems(I_COL_SOUM_PROFIT).Tag, 2) Then
2130                      bExiste = True

2135                      Exit For
2140                    End If
2145                  Next

2150                  If bExiste = False Then
2155                    Call collExtra.Add(Right$(itmPiece.ListSubItems(I_COL_SOUM_PROFIT).Tag, 2))
2160                  End If
2165                Else
2170                  If Left$(itmPiece.ListSubItems(I_COL_SOUM_PROFIT).Tag, 10) = "ANNULATION" Then
2175                    Call AjouterAnnulationDansExtra(itmPiece, sExtra)

2180                    bCalculExtra = True

2185                    bExiste = False
  
2190                    For iCompteurExtra = 1 To collExtra.count
2195                      If collExtra(iCompteurExtra) = Right$(itmPiece.ListSubItems(I_COL_SOUM_PROFIT).Tag, 2) Then
2200                        bExiste = True

2205                        Exit For
2210                      End If
2215                    Next

2220                    If bExiste = False Then
2225                      Call collExtra.Add(Right$(itmPiece.ListSubItems(I_COL_SOUM_PROFIT).Tag, 2))
2230                    End If
2235                  End If
2240                End If
2245              End If
2250            End If
2255          End If
2260        End If
2265      End If
2270    Next

2275    If m_eType = TYPE_PROJET Then
2280      If Right$(txtNoProjSoum.Text, 2) <= 98 And Right$(txtNoProjSoum.Text, 2) >= 60 Then
2285        Call CalculerTotalRecordset(Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & rstProjSoum.Fields("LiaisonChargeable"))
2290      End If
2295    End If

2300    If bCalculExtra = True Then
2305      For iCompteurExtra = 1 To collExtra.count
2310        Call CalculerTotalRecordset(Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & collExtra(iCompteurExtra))
2315      Next
2320    End If
     
2325    rstProjSoum.Fields("total_temps") = txtTotalTemps.Text

2330    Call rstProjSoum.Update

2335    Call rstProjSoum.Close
2340    Set rstProjSoum = Nothing

2345    Call rstPiece.Close
2350    Set rstPiece = Nothing

2355    If m_eType = TYPE_SOUMISSION Then
2360      Call AjouterSoumissionAuCumulatif
2365    Else
2370      Call AjouterProjetAuCumulatif
2375    End If

2380    Exit Sub

AfficherErreur:

2385    woups "frmProjSoumElec", "EnregistrerProjSoum", Err, Erl, sNoProjSoum)

  'Si un erreur se produit dans l'enregistrement des pièces, il faut avertir
  'l'utilisateur de quelle pièce il s'agit et continuer avec un Resume Next
2390    If Erl >= 1310 And Erl <= 2265 Then
2395      If m_eLangage = ANGLAIS Then
2400        Call rstSection.Open("SELECT NomSectionEN FROM GRB_SoumProjSectionElec", g_connData, adOpenDynamic, adLockOptimistic)
      
2405        If Not rstSection.EOF Then
2410          sSection = rstSection.Fields("NomSectionEN")
2415        Else
2420          sSection = itmPiece.Tag
2425        End If
2430      Else
2435        Call rstSection.Open("SELECT NomSectionFR FROM GRB_SoumProjSectionElec", g_connData, adOpenDynamic, adLockOptimistic)
      
2440        If Not rstSection.EOF Then
2445          sSection = rstSection.Fields("NomSectionFR")
2450        Else
2455          sSection = itmPiece.Tag
2460        End If
2465      End If
     
2470      Call rstSection.Close
2475      Set rstSection = Nothing
  
2480      Call MsgBox("La pièce " & itmPiece.SubItems(I_COL_SOUM_PIECE) & " de la section " & sSection & " risque de contenir des erreurs." & vbNewLine & _
                      "Il se peut qu'elle ne soit plus présente dans la liste.")
2485    End If
   
2490    Resume Next
End Sub

Private Sub InitialiserNouveauxTaux()

5       On Error GoTo AfficherErreur

10      Dim rstConfig As ADODB.Recordset

15      Set rstConfig = New ADODB.Recordset

20      Call rstConfig.Open("SELECT TauxDessinElec, TauxFabrication, TauxAssemblageElec, TauxProgInterface, TauxProgAutomate, TauxProgRobot, TauxVision, TauxTestElec, TauxInstallationElec, TauxMiseService, TauxFormationElec, TauxGestionProjetsElec, TauxShippingElec, Hebergement1, Hebergement2, Repas, Standard, UniteMobile FROM GRB_Config", g_connData, adOpenDynamic, adLockOptimistic)

25      If Not IsNull(rstConfig.Fields("TauxDessinElec")) Then
30        m_sTauxDessin = rstConfig.Fields("TauxDessinElec")
35      Else
40        m_sTauxDessin = "0"
45      End If

50      If Not IsNull(rstConfig.Fields("TauxFabrication")) Then
55        m_sTauxFabrication = rstConfig.Fields("TauxFabrication")
60      Else
65        m_sTauxFabrication = "0"
70      End If

75      If Not IsNull(rstConfig.Fields("TauxAssemblageElec")) Then
80        m_sTauxAssemblage = rstConfig.Fields("TauxAssemblageElec")
85      Else
90        m_sTauxAssemblage = "0"
95      End If

100     If Not IsNull(rstConfig.Fields("TauxProgInterface")) Then
105       m_sTauxProgInterface = rstConfig.Fields("TauxProgInterface")
110     Else
115       m_sTauxProgInterface = "0"
120     End If

125     If Not IsNull(rstConfig.Fields("TauxProgAutomate")) Then
130       m_sTauxProgAutomate = rstConfig.Fields("TauxProgAutomate")
135     Else
140       m_sTauxProgAutomate = "0"
145     End If

150     If Not IsNull(rstConfig.Fields("TauxProgRobot")) Then
155       m_sTauxProgRobot = rstConfig.Fields("TauxProgRobot")
160     Else
165       m_sTauxProgRobot = "0"
170     End If

175     If Not IsNull(rstConfig.Fields("TauxVision")) Then
180       m_sTauxVision = rstConfig.Fields("TauxVision")
185     Else
190       m_sTauxVision = "0"
195     End If

200     If Not IsNull(rstConfig.Fields("TauxTestElec")) Then
205       m_sTauxTest = rstConfig.Fields("TauxTestElec")
210     Else
215       m_sTauxTest = "0"
220     End If

225     If Not IsNull(rstConfig.Fields("TauxInstallationElec")) Then
230       m_sTauxInstallation = rstConfig.Fields("TauxInstallationElec")
235     Else
240       m_sTauxInstallation = "0"
245     End If

250     If Not IsNull(rstConfig.Fields("TauxMiseService")) Then
255       m_sTauxMiseService = rstConfig.Fields("TauxMiseService")
260     Else
265       m_sTauxMiseService = "0"
270     End If

275     If Not IsNull(rstConfig.Fields("TauxFormationElec")) Then
280       m_sTauxFormation = rstConfig.Fields("TauxFormationElec")
285     Else
290       m_sTauxFormation = "0"
295     End If

300     If Not IsNull(rstConfig.Fields("TauxGestionProjetsElec")) Then
305       m_sTauxGestion = rstConfig.Fields("TauxGestionProjetsElec")
310     Else
315       m_sTauxGestion = "0"
320     End If

325     If Not IsNull(rstConfig.Fields("TauxShippingElec")) Then
330       m_sTauxShipping = rstConfig.Fields("TauxShippingElec")
335     Else
340       m_sTauxShipping = "0"
345     End If

350     If m_eType = TYPE_PROJET Then
355       m_sTauxHebergement1 = "0"
360       m_sTauxHebergement2 = "0"
365       m_sTauxRepas = "0"
370       m_sTauxTransport = "0"
375       m_sTauxUniteMobile = "0"
380     Else
385       m_sTauxHebergement1 = rstConfig.Fields("Hebergement1")
390       m_sTauxHebergement2 = rstConfig.Fields("Hebergement2")
395       m_sTauxRepas = rstConfig.Fields("Repas")
400       m_sTauxTransport = rstConfig.Fields("Standard")
405       m_sTauxUniteMobile = rstConfig.Fields("UniteMobile")
410     End If

415     Call rstConfig.Close
420     Set rstConfig = Nothing

425     Exit Sub

AfficherErreur:

430     woups "frmProjSoumElec", "InitialiserNouveauxTaux", Err, Erl
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
        
45      Call rstProjet.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sLiaison & "'", g_connData, adOpenDynamic, adLockOptimistic)

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
       
110       rstPiece.Fields("Type") = "E"
       
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
        
215       rstPiece.Fields("Temps") = Trim$(itmSource.SubItems(I_COL_SOUM_TEMPS))
220       rstPiece.Fields("Temps_Total") = itmSource.SubItems(I_COL_SOUM_MONTAGE)
225       rstPiece.Fields("Prix_Total") = Conversion(itmSource.SubItems(I_COL_SOUM_TOTAL), MODE_PAS_FORMAT)
230       rstPiece.Fields("Profit_argent") = Conversion(itmSource.SubItems(I_COL_SOUM_PROFIT), MODE_PAS_FORMAT)
235       rstPiece.Fields("SousSection") = itmSource.ListSubItems(I_COL_SOUM_PIECE).Tag
            
240       If itmSource.SubItems(I_COL_SOUM_PRIX_LIST) <> vbNullString Then
245         If itmSource.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag <> vbNullString Then
250           rstPiece.Fields("PrixOrigine") = Replace(Round(CDbl(Replace(itmSource.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag, ".", ",")), 4), ".", ",")
255         Else
260           rstPiece.Fields("PrixOrigine") = "0"
265         End If
270       Else
275         rstPiece.Fields("PrixOrigine") = "0"
280       End If
    
285       If InStr(1, itmSource.Text, "*") > 0 Then
290         rstPiece.Fields("Quoté") = True
295       Else
300         rstPiece.Fields("Quoté") = False
305       End If

310       If Trim$(itmSource.SubItems(I_COL_SOUM_ID) <> vbNullString) Then
315         rstPiece.Fields("ID") = itmSource.SubItems(I_COL_SOUM_ID)
320       End If

325       rstPiece.Fields("PieceExtraChargeable") = True
330       rstPiece.Fields("Provenance") = Right$(txtNoProjSoum.Text, 2)

335       Call rstPiece.Update

340       Call rstPiece.Close

345       rstPiece.CursorLocation = adUseServer

350       Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sLiaison & "' AND NuméroLigne >= " & iCompteur & " ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)

          'Tant qu'il y a des enregistrements dans le recordset
355       Do While Not rstPiece.EOF
360         If rstPiece.Fields("PieceExtraChargeable") = True And rstPiece.Fields("Qté") = itmSource.Text And rstPiece.Fields("NumItem") = itmSource.SubItems(I_COL_SOUM_PIECE) And bSkip = False Then
365           bSkip = True
370         Else
375           rstPiece.Fields("NuméroLigne") = rstPiece.Fields("NuméroLigne") + 1

380           Call rstPiece.Update
385         End If

390         Call rstPiece.MoveNext
395       Loop

400       Call rstPiece.Close
405       Set rstPiece = Nothing

410       Call CalculerTempsFabricationRecordset(Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sLiaison)
415     End If

420     Call rstProjet.Close
425     Set rstProjet = Nothing

430     Exit Sub

AfficherErreur:

435     woups "frmProjSoumElec", "AjouterPiecesExtraDansJob", Err, Erl

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
        
45      Call rstProjet.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sLiaison & "'", g_connData, adOpenDynamic, adLockOptimistic)

        'Si le projet existe
50      If Not rstProjet.EOF Then
          'Ouverture du recordset sur le projet original
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
       
105       rstPiece.Fields("Type") = "E"
       
110       If itmSource.Checked = True Then
115         rstPiece.Fields("Visible") = True
120       Else
125         rstPiece.Fields("Visible") = False
130       End If

135       rstPiece.Fields("Facturation") = itmSource.SubItems(I_COL_SOUM_FACTURATION)
              
140       rstPiece.Fields("IDSection") = itmSource.Tag
145       rstPiece.Fields("NumItem") = Trim$(itmSource.SubItems(I_COL_SOUM_PIECE))
150       rstPiece.Fields("Qté") = Replace(itmSource.Text, "*", vbNullString)
155       rstPiece.Fields("Desc_FR") = Trim$(itmSource.SubItems(I_COL_SOUM_DESCR))
160       rstPiece.Fields("Desc_EN") = Trim$(itmSource.ListSubItems(I_COL_SOUM_DESCR).Tag)
165       rstPiece.Fields("Manufact") = Trim$(itmSource.SubItems(I_COL_SOUM_MANUFACT))
170       rstPiece.Fields("Prix_list") = Conversion(itmSource.SubItems(I_COL_SOUM_PRIX_LIST), MODE_PAS_FORMAT, 4)
175       rstPiece.Fields("Escompte") = Conversion(itmSource.SubItems(I_COL_SOUM_ESCOMPTE), MODE_PAS_FORMAT)
180       rstPiece.Fields("Prix_net") = Conversion(itmSource.SubItems(I_COL_SOUM_PRIX_NET), MODE_PAS_FORMAT, 4)
185       rstPiece.Fields("OrdreSection") = itmSource.ListSubItems(I_COL_SOUM_MANUFACT).Tag
190       rstPiece.Fields("NuméroLigne") = iCompteur
      
195       If itmSource.SubItems(I_COL_SOUM_DISTRIB) <> vbNullString Then
200         rstPiece.Fields("IDFRS") = itmSource.ListSubItems(I_COL_SOUM_DISTRIB).Tag
205       End If
        
210       rstPiece.Fields("Temps") = Trim$(itmSource.SubItems(I_COL_SOUM_TEMPS))
215       rstPiece.Fields("Temps_Total") = itmSource.SubItems(I_COL_SOUM_MONTAGE)
220       rstPiece.Fields("Prix_Total") = Conversion(itmSource.SubItems(I_COL_SOUM_TOTAL), MODE_PAS_FORMAT)
225       rstPiece.Fields("Profit_argent") = Conversion(itmSource.SubItems(I_COL_SOUM_PROFIT), MODE_PAS_FORMAT)
230       rstPiece.Fields("SousSection") = itmSource.ListSubItems(I_COL_SOUM_PIECE).Tag
            
235       If itmSource.SubItems(I_COL_SOUM_PRIX_LIST) <> vbNullString Then
240         If itmSource.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag <> vbNullString Then
245           rstPiece.Fields("PrixOrigine") = Replace(Round(CDbl(Replace(itmSource.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag, ".", ",")), 2), ".", ",")
250         Else
255           rstPiece.Fields("PrixOrigine") = "0"
260         End If
265       Else
270         rstPiece.Fields("PrixOrigine") = "0"
275       End If
    
280       If InStr(1, itmSource.Text, "*") > 0 Then
285         rstPiece.Fields("Quoté") = True
290       Else
295         rstPiece.Fields("Quoté") = False
300       End If

305       If Trim$(itmSource.SubItems(I_COL_SOUM_ID) <> vbNullString) Then
310         rstPiece.Fields("ID") = itmSource.SubItems(I_COL_SOUM_ID)
315       End If

320       rstPiece.Fields("PieceExtraNonChargeable") = True
325       rstPiece.Fields("Provenance") = Right$(txtNoProjSoum.Text, 2)

330       Call rstPiece.Update

335       Call rstPiece.Close

340       rstPiece.CursorLocation = adUseServer

345       Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sLiaison & "' AND NuméroLigne >= " & iCompteur & " ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)

          'Tant qu'il y a des enregistrements dans le recordset
350       Do While Not rstPiece.EOF
355         If rstPiece.Fields("PieceExtraNonChargeable") = True And rstPiece.Fields("Qté") = itmSource.Text And rstPiece.Fields("NumItem") = itmSource.SubItems(I_COL_SOUM_PIECE) And bSkip = False Then
360           bSkip = True
365         Else
370           rstPiece.Fields("NuméroLigne") = rstPiece.Fields("NuméroLigne") + 1

375           Call rstPiece.Update
380         End If

385         Call rstPiece.MoveNext
390       Loop

395       Call rstPiece.Close
400       Set rstPiece = Nothing

405       Call CalculerTempsFabricationRecordset(Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sLiaison)
410     End If

415     Call rstProjet.Close
420     Set rstProjet = Nothing

425     Exit Sub

AfficherErreur:

430     woups "frmProjSoumElec", "AjouterPiecesExtraDansJob", Err, Erl

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
  
  Call MsgBox("La pièce " & itmSource.SubItems(I_COL_SOUM_PIECE) & " de la section " & sSection & " risque de contenir des erreurs dans le projet " & sLiaison & "." & vbNewLine & _
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
        
45      Call rstProjet.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sExtra & "'", g_connData, adOpenDynamic, adLockOptimistic)

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
       
105       rstPiece.Fields("Type") = "E"
       
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
     
255       rstPiece.Fields("Temps") = Trim$(itmSource.SubItems(I_COL_SOUM_TEMPS))
260       rstPiece.Fields("Temps_Total") = itmSource.SubItems(I_COL_SOUM_MONTAGE)
265       rstPiece.Fields("Prix_Total") = Conversion(itmSource.SubItems(I_COL_SOUM_TOTAL), MODE_PAS_FORMAT)
270       rstPiece.Fields("Profit_argent") = Conversion(itmSource.SubItems(I_COL_SOUM_PROFIT), MODE_PAS_FORMAT)
275       rstPiece.Fields("SousSection") = itmSource.ListSubItems(I_COL_SOUM_PIECE).Tag
            
280       If itmSource.SubItems(I_COL_SOUM_PRIX_LIST) <> vbNullString Then
285         If itmSource.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag <> vbNullString Then
290           rstPiece.Fields("PrixOrigine") = Replace(Round(CDbl(Replace(itmSource.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag, ".", ",")), 4), ".", ",")
295         Else
300           rstPiece.Fields("PrixOrigine") = "0"
305         End If
310       Else
315         rstPiece.Fields("PrixOrigine") = "0"
320       End If
  
325       If InStr(1, itmSource.Text, "*") > 0 Then
330         rstPiece.Fields("Quoté") = True
335       Else
340         rstPiece.Fields("Quoté") = False
345       End If

350       rstPiece.Fields("DateRetour") = itmSource.ListSubItems(I_COL_SOUM_DATE_REQUISE).Tag

355       rstPiece.Fields("Commentaire") = itmSource.SubItems(I_COL_SOUM_COMMENTAIRE)

360       Call rstPiece.Update

365       Call rstPiece.Close

370       rstPiece.CursorLocation = adUseServer

375       Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sExtra & "' AND NuméroLigne >= " & iCompteur & " ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)

          'Tant qu'il y a des enregistrements dans le recordset
380       Do While Not rstPiece.EOF
385         If itmSource.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_BRUN Then
390           If rstPiece.Fields("MatérielInutile") = True And rstPiece.Fields("Qté") = itmSource.Text And rstPiece.Fields("NumItem") = itmSource.SubItems(I_COL_SOUM_PIECE) And bSkip = False Then
395             bSkip = True
400           Else
405             rstPiece.Fields("NuméroLigne") = rstPiece.Fields("NuméroLigne") + 1

410             Call rstPiece.Update
415           End If
420         Else
425           If rstPiece.Fields("MatérielInutile") = False And rstPiece.Fields("Qté") = itmSource.Text And rstPiece.Fields("NumItem") = itmSource.SubItems(I_COL_SOUM_PIECE) And bSkip = False Then
430             bSkip = True
435           Else
440             rstPiece.Fields("NuméroLigne") = rstPiece.Fields("NuméroLigne") + 1

445             Call rstPiece.Update
450           End If
455         End If

460         Call rstPiece.MoveNext
465       Loop

470       Call rstPiece.Close
475       Set rstPiece = Nothing

480       Call CalculerTempsFabricationRecordset(Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sExtra)
485     End If

490     Call rstProjet.Close
495     Set rstProjet = Nothing

500     Exit Sub

AfficherErreur:

505     woups "frmProjSoumElec", "AjouterInutileDansExtra", Err, Erl

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
        
45      Call rstProjet.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sExtra & "'", g_connData, adOpenDynamic, adLockOptimistic)

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
       
105       rstPiece.Fields("Type") = "E"
       
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
     
235       rstPiece.Fields("Temps") = Trim$(itmSource.SubItems(I_COL_SOUM_TEMPS))
240       rstPiece.Fields("Temps_Total") = itmSource.SubItems(I_COL_SOUM_MONTAGE)
245       rstPiece.Fields("Prix_Total") = Conversion(itmSource.SubItems(I_COL_SOUM_TOTAL), MODE_PAS_FORMAT)
250       rstPiece.Fields("Profit_argent") = Conversion(itmSource.SubItems(I_COL_SOUM_PROFIT), MODE_PAS_FORMAT)
255       rstPiece.Fields("SousSection") = itmSource.ListSubItems(I_COL_SOUM_PIECE).Tag
            
260       If itmSource.SubItems(I_COL_SOUM_PRIX_LIST) <> vbNullString Then
265         If itmSource.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag <> vbNullString Then
270           rstPiece.Fields("PrixOrigine") = Replace(Round(CDbl(Replace(itmSource.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag, ".", ",")), 2), ".", ",")
275         Else
280           rstPiece.Fields("PrixOrigine") = "0"
285         End If
290       Else
295         rstPiece.Fields("PrixOrigine") = "0"
300       End If
  
305       If InStr(1, itmSource.Text, "*") > 0 Then
310         rstPiece.Fields("Quoté") = True
315       Else
320         rstPiece.Fields("Quoté") = False
325       End If

330       rstPiece.Fields("Commentaire") = itmSource.SubItems(I_COL_SOUM_COMMENTAIRE)

335       Call rstPiece.Update

340       Call rstPiece.Close

345       rstPiece.CursorLocation = adUseServer

350       Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sExtra & "' AND NuméroLigne >= " & iCompteur & " ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)

          'Tant qu'il y a des enregistrements dans le recordset
355       Do While Not rstPiece.EOF
360         If rstPiece.Fields("CommandeAnnulée") = True And rstPiece.Fields("Qté") = itmSource.Text And rstPiece.Fields("NumItem") = itmSource.SubItems(I_COL_SOUM_PIECE) And bSkip = False Then
365           bSkip = True
370         Else
375           rstPiece.Fields("NuméroLigne") = rstPiece.Fields("NuméroLigne") + 1

380           Call rstPiece.Update
385         End If

390         Call rstPiece.MoveNext
395       Loop

400       Call rstPiece.Close
405       Set rstPiece = Nothing

410       Call CalculerTempsFabricationRecordset(Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sExtra)
415     End If

420     Call rstProjet.Close
425     Set rstProjet = Nothing

430     Exit Sub

AfficherErreur:

435     woups "frmProjSoumElec", "AjouterAnnulationDansExtra", Err, Erl

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

25      woups "frmProjSoumElec", "cmdFermer_Click", Err, Erl
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
    
45      Set rstProjSoum = New ADODB.Recordset
50      Set rstEmploye = New ADODB.Recordset
55      Set rstCreation = New ADODB.Recordset
    
        'Il faut le vider avant de le remplir
60      Call lvwHistorique.ListItems.Clear
          
65      If m_eType = TYPE_PROJET Then
70        sChamps = "IDProjet"
75        sTable = "GRB_Projet_Modif"
80        sTableCreer = "GRB_ProjetElec"
85      Else
90        sChamps = "IDSoumission"
95        sTable = "GRB_Soumission_Modif"
100       sTableCreer = "GRB_SoumissionElec"
105     End If
  
        'Ouverture du recordset selon le type
110     Call rstCreation.Open("SELECT creer, creer_par FROM " & sTableCreer & " WHERE " & sChamps & " = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
  
        'Ajout de la section "Création"
115     Set itmModif = lvwHistorique.ListItems.Add
    
120     itmModif.Text = "Création"
  
125     itmModif.Bold = True
   
        'Ajout du nom de celui qui l'a créé
130     Set itmModif = lvwHistorique.ListItems.Add
  
135     Call rstEmploye.Open("SELECT Employe FROM GRB_Employés WHERE noEmploye = " & rstCreation.Fields("creer_par"), g_connData, adOpenDynamic, adLockOptimistic)
  
140     itmModif.Text = rstEmploye.Fields("Employe")
 
145     Call rstEmploye.Close
  
        'Date
150     itmModif.SubItems(I_COL_MODIF_DATE) = rstCreation.Fields("creer")
  
155     itmModif.SubItems(I_COL_MODIF_HEURE) = vbNullString
  
160     Call rstCreation.Close
165     Set rstCreation = Nothing
  
170     Call rstProjSoum.Open("SELECT * FROM " & sTable & " WHERE " & sChamps & " = '" & txtNoProjSoum.Text & "' AND Type = 'E' AND TypeModif = 'MODIFICATION' ORDER BY [Date], Heure", g_connData, adOpenDynamic, adLockOptimistic)
             
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

250     Call rstProjSoum.Open("SELECT * FROM " & sTable & " WHERE " & sChamps & " = '" & txtNoProjSoum.Text & "' AND Type = 'E' AND TypeModif = 'RECEPTION' ORDER BY [Date], Heure", g_connData, adOpenDynamic, adLockOptimistic)
             
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

330     Call rstProjSoum.Open("SELECT * FROM " & sTable & " WHERE " & sChamps & " = '" & txtNoProjSoum.Text & "' AND Type = 'E' AND TypeModif = 'RETOUR' ORDER BY [Date], Heure", g_connData, adOpenDynamic, adLockOptimistic)
             
335     If Not rstProjSoum.EOF Then
          'Ajout de la section "Retour de marchandise"
340       Set itmModif = lvwHistorique.ListItems.Add
    
345       itmModif.Text = "Retour de marchandise"
    
350       itmModif.Bold = True
    
355       Do While Not rstProjSoum.EOF
            'Ajout des retours de marchandise
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
 
410     Call rstProjSoum.Open("SELECT * FROM " & sTable & " WHERE " & sChamps & " = '" & txtNoProjSoum.Text & "' AND Type = 'E' AND TypeModif = 'FACTURATION' ORDER BY [Date], Heure", g_connData, adOpenDynamic, adLockOptimistic)
             
415     If Not rstProjSoum.EOF Then
          'Ajout de la section "Modifications"
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

510     woups "frmProjSoumElec", "RemplirListViewModifications", Err, Erl
End Sub

Private Sub RemplirListViewSuppression()

5       On Error GoTo AfficherErreur

        'Rempli le listView avec les pièces supprimées
10      Dim rstBavard  As ADODB.Recordset
15      Dim rstEmploye As ADODB.Recordset
20      Dim itmBavard  As ListItem

25      Call lvwBavard.ListItems.Clear

30      Set rstBavard = New ADODB.Recordset
35      Set rstEmploye = New ADODB.Recordset

40      Call rstBavard.Open("SELECT * FROM GRB_BavardSuppression WHERE NoProjSoum = '" & txtNoProjSoum.Text & "' AND Type = 'E' ORDER BY [Date], Heure", g_connData, adOpenDynamic, adLockOptimistic)

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

100     Call rstBavard.Close
105     Set rstBavard = Nothing

110     Set rstEmploye = Nothing

115     Exit Sub

AfficherErreur:

120     woups "frmProjSoumElec", "RemplirListViewSuppression", Err, Erl
End Sub

Private Sub Cmdajouter_Click()

5       On Error GoTo AfficherErreur

        'Ajoute une soumission ou un projet
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
180         If ValiderFormatElectrique(sNumero) = False Then
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
285       Call rstProjSoum.Open("SELECT * FROM GRB_SoumissionElec WHERE IDSoumission = '" & sNumero & "'", g_connData, adOpenDynamic, adLockPessimistic)
        
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

485         If Not rstProjSoum.EOF Then
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
    
565           Call InitialiserVariables(txtNoProjSoum.Text)
    
              'Débarre les champs
570           Call BarrerChamps(False)
        
              'Vide les champs
575           Call ViderChamps
580         Else
585           If VerifierProjet(sNoProjet) = True Then
                'Débarre les champs
590             Call BarrerChamps(False)
                
                'Vide les champs
595             Call ViderChamps

600             txtNoProjSoum.Text = sNumero

605             Call RemplirSoumissionProjet(sNumero, sNoProjet)
610           Else
615             Call MsgBox("Le projet " & sNoProjet & " n'existe pas!", vbOKOnly, "Erreur")

620             Screen.MousePointer = vbDefault

625             Exit Sub
630           End If
635         End If

            'Vide la valeur par défaut si demande Sous-Section
640         m_sSousSection = vbNullString
        
645         m_bModeAjout = True
650         m_bModeAffichage = False
                
655         lvwSoumission.Height = lvwSoumission.Height * 0.49
660         lvwSoumission.Top = lvwPieces.Top + lvwPieces.Height + (lvwSoumission.Height * 0.02)
        
            'Met le form en mode ajout/modif
665         Call AfficherControles(MODE_AJOUT_MODIF)
670       End If
675     End If
  
680     Screen.MousePointer = vbDefault

685     Exit Sub

AfficherErreur:

690     woups "frmProjSoumElec", "cmdAjouter_Click", Err, Erl
End Sub

Private Function VerifierProjet(ByVal sNoProjet As String) As Boolean
  
5       On Error GoTo AfficherErreur

10      Dim rstProjet As ADODB.Recordset
15      Dim bExiste   As Boolean

20      Set rstProjet = New ADODB.Recordset

25      Call rstProjet.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)

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

75      woups "frmProjSoumElec", "VerifierProjet", Err, Erl
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
                  "-  % Imprévu" & vbNewLine & _
                  "-  $ Pages manuel", vbYesNo) = vbYes Then
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
    
120     Call rstProjSoum.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)

125     m_bSansTemps = rstProjSoum.Fields("SansTemps")

130     If Not IsNull(rstProjSoum.Fields("TempsDessin")) Then
135       m_sTempsDessin = rstProjSoum.Fields("TempsDessin")
140     Else
145       m_sTempsDessin = "0"
150     End If

155     If Not IsNull(rstProjSoum.Fields("TempsFabrication")) Then
160       m_sTempsFabrication = rstProjSoum.Fields("TempsFabrication")
165     Else
170       m_sTempsFabrication = "0"
175     End If

180     If Not IsNull(rstProjSoum.Fields("TempsAssemblage")) Then
185       m_sTempsAssemblage = rstProjSoum.Fields("TempsAssemblage")
190     Else
195       m_sTempsAssemblage = "0"
200     End If

205     If Not IsNull(rstProjSoum.Fields("TempsProgInterface")) Then
210       m_sTempsProgInterface = rstProjSoum.Fields("TempsProgInterface")
215     Else
220       m_sTempsProgInterface = "0"
225     End If

230     If Not IsNull(rstProjSoum.Fields("TempsProgAutomate")) Then
235       m_sTempsProgAutomate = rstProjSoum.Fields("TempsProgAutomate")
240     Else
245       m_sTempsProgAutomate = "0"
250     End If

255     If Not IsNull(rstProjSoum.Fields("TempsProgRobot")) Then
260       m_sTempsProgRobot = rstProjSoum.Fields("TempsProgRobot")
265     Else
270       m_sTempsProgRobot = "0"
275     End If

280     If Not IsNull(rstProjSoum.Fields("TempsVision")) Then
285       m_sTempsVision = rstProjSoum.Fields("TempsVision")
290     Else
295       m_sTempsVision = "0"
300     End If

305     If Not IsNull(rstProjSoum.Fields("TempsTest")) Then
310       m_sTempsTest = rstProjSoum.Fields("TempsTest")
315     Else
320       m_sTempsTest = "0"
325     End If

330     If Not IsNull(rstProjSoum.Fields("TempsInstallation")) Then
335       m_sTempsInstallation = rstProjSoum.Fields("TempsInstallation")
340     Else
345       m_sTempsInstallation = "0"
350     End If

355     If Not IsNull(rstProjSoum.Fields("TempsMiseService")) Then
360       m_sTempsMiseService = rstProjSoum.Fields("TempsMiseService")
365     Else
370       m_sTempsMiseService = "0"
375     End If

380     If Not IsNull(rstProjSoum.Fields("TempsFormation")) Then
385       m_sTempsFormation = rstProjSoum.Fields("TempsFormation")
390     Else
395       m_sTempsFormation = "0"
400     End If

405     If Not IsNull(rstProjSoum.Fields("TempsGestion")) Then
410       m_sTempsGestion = rstProjSoum.Fields("TempsGestion")
415     Else
420       m_sTempsGestion = "0"
425     End If

430     If Not IsNull(rstProjSoum.Fields("TempsShipping")) Then
435       m_sTempsShipping = rstProjSoum.Fields("TempsShipping")
440     Else
445       m_sTempsShipping = "0"
450     End If
  
455     Call rstConfig.Open("SELECT * FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

460     If bTauxHoraire = True Then
465       If Not IsNull(rstConfig.Fields("TauxDessinElec")) Then
470         m_sTauxDessin = rstConfig.Fields("TauxDessinElec")
475       Else
480         m_sTauxDessin = "0"
485       End If

490       If Not IsNull(rstConfig.Fields("TauxFabrication")) Then
495         m_sTauxFabrication = rstConfig.Fields("TauxFabrication")
500       Else
505         m_sTauxFabrication = "0"
510       End If

515       If Not IsNull(rstConfig.Fields("TauxAssemblageElec")) Then
520         m_sTauxAssemblage = rstConfig.Fields("TauxAssemblageElec")
525       Else
530         m_sTauxAssemblage = "0"
535       End If

540       If Not IsNull(rstConfig.Fields("TauxProgInterface")) Then
545         m_sTauxProgInterface = rstConfig.Fields("TauxProgInterface")
550       Else
555         m_sTauxProgInterface = "0"
560       End If

565       If Not IsNull(rstConfig.Fields("TauxProgAutomate")) Then
570         m_sTauxProgAutomate = rstConfig.Fields("TauxProgAutomate")
575       Else
580         m_sTauxProgAutomate = "0"
585       End If

590       If Not IsNull(rstConfig.Fields("TauxProgRobot")) Then
595         m_sTauxProgRobot = rstConfig.Fields("TauxProgRobot")
600       Else
605         m_sTauxProgRobot = "0"
610       End If

615       If Not IsNull(rstConfig.Fields("TauxVision")) Then
620         m_sTauxVision = rstConfig.Fields("TauxVision")
625       Else
630         m_sTauxVision = "0"
635       End If

640       If Not IsNull(rstConfig.Fields("TauxTestElec")) Then
645         m_sTauxTest = rstConfig.Fields("TauxTestElec")
650       Else
655         m_sTauxTest = "0"
660       End If

665       If Not IsNull(rstConfig.Fields("TauxInstallationElec")) Then
670         m_sTauxInstallation = rstConfig.Fields("TauxInstallationElec")
675       Else
680         m_sTauxInstallation = "0"
685       End If

690       If Not IsNull(rstConfig.Fields("TauxMiseService")) Then
695         m_sTauxMiseService = rstConfig.Fields("TauxMiseService")
700       Else
705         m_sTauxMiseService = "0"
710       End If

715       If Not IsNull(rstConfig.Fields("TauxFormationElec")) Then
720         m_sTauxFormation = rstConfig.Fields("TauxFormationElec")
725       Else
730         m_sTauxFormation = "0"
735       End If

740       If Not IsNull(rstConfig.Fields("TauxGestionProjetsElec")) Then
745         m_sTauxGestion = rstConfig.Fields("TauxGestionProjetsElec")
750       Else
755         m_sTauxGestion = "0"
760       End If

765       If Not IsNull(rstConfig.Fields("TauxShippingElec")) Then
770         m_sTauxShipping = rstConfig.Fields("TauxShippingElec")
775       Else
780         m_sTauxShipping = "0"
785       End If
790     Else
795       If Not IsNull(rstProjSoum.Fields("TauxDessin")) Then
800         m_sTauxDessin = rstProjSoum.Fields("TauxDessin")
805       Else
810         m_sTauxDessin = "0"
815       End If

820       If Not IsNull(rstProjSoum.Fields("TauxFabrication")) Then
825         m_sTauxFabrication = rstProjSoum.Fields("TauxFabrication")
830       Else
835         m_sTauxFabrication = "0"
840       End If

845       If Not IsNull(rstProjSoum.Fields("TauxAssemblage")) Then
850         m_sTauxAssemblage = rstProjSoum.Fields("TauxAssemblage")
855       Else
860         m_sTauxAssemblage = "0"
865       End If

870       If Not IsNull(rstProjSoum.Fields("TauxProgInterface")) Then
875         m_sTauxProgInterface = rstProjSoum.Fields("TauxProgInterface")
880       Else
885         m_sTauxProgInterface = "0"
890       End If

895       If Not IsNull(rstProjSoum.Fields("TauxProgAutomate")) Then
900         m_sTauxProgAutomate = rstProjSoum.Fields("TauxProgAutomate")
905       Else
910         m_sTauxProgAutomate = "0"
915       End If

920       If Not IsNull(rstProjSoum.Fields("TauxProgRobot")) Then
925         m_sTauxProgRobot = rstProjSoum.Fields("TauxProgRobot")
930       Else
935         m_sTauxProgRobot = "0"
940       End If

945       If Not IsNull(rstProjSoum.Fields("TauxVision")) Then
950         m_sTauxVision = rstProjSoum.Fields("TauxVision")
955       Else
960         m_sTauxVision = "0"
965       End If

970       If Not IsNull(rstProjSoum.Fields("TauxTest")) Then
975         m_sTauxTest = rstProjSoum.Fields("TauxTest")
980       Else
985         m_sTauxTest = "0"
990       End If

995       If Not IsNull(rstProjSoum.Fields("TauxInstallation")) Then
1000        m_sTauxInstallation = rstProjSoum.Fields("TauxInstallation")
1005      Else
1010        m_sTauxInstallation = "0"
1015      End If

1020      If Not IsNull(rstProjSoum.Fields("TauxMiseService")) Then
1025        m_sTauxMiseService = rstProjSoum.Fields("TauxMiseService")
1030      Else
1035        m_sTauxMiseService = "0"
1040      End If

1045      If Not IsNull(rstProjSoum.Fields("TauxFormation")) Then
1050        m_sTauxFormation = rstProjSoum.Fields("TauxFormation")
1055      Else
1060        m_sTauxFormation = "0"
1065      End If

1070      If Not IsNull(rstProjSoum.Fields("TauxGestion")) Then
1075        m_sTauxGestion = rstProjSoum.Fields("TauxGestion")
1080      Else
1085        m_sTauxGestion = "0"
1090      End If

1095      If Not IsNull(rstProjSoum.Fields("TauxShipping")) Then
1100        m_sTauxShipping = rstProjSoum.Fields("TauxShipping")
1105      Else
1110        m_sTauxShipping = "0"
1115      End If
1120    End If

1125    If bVariables = True Then
1130      m_sProfit = rstConfig.Fields("ProfitElec")
1135      m_sCommission = rstConfig.Fields("Commission")
1140      m_sImprevue = rstConfig.Fields("Imprévus")
1145    Else
1150      m_sProfit = rstProjSoum.Fields("Profit")
1155      m_sCommission = rstProjSoum.Fields("Commission")
1160      m_sImprevue = rstProjSoum.Fields("Imprevue")
1165    End If

1170    Call rstConfig.Close
1175    Set rstConfig = Nothing
          
1180    txtProjet.Text = rstProjSoum.Fields("Description")
1185    txtNbreManuel.Text = rstProjSoum.Fields("NbreManuel")
1190    txtPrixManuel.Text = rstProjSoum.Fields("total_manuel")
1195    txtTransport.Text = rstProjSoum.Fields("transport")

1200    If Not IsNull(rstProjSoum.Fields("CheminPhotos")) Then
1205      txtCheminPhotos.Text = rstProjSoum.Fields("CheminPhotos")
1210    Else
1215      txtCheminPhotos.Text = vbNullString
1220    End If
  
1225    chkCSA.Value = Abs(CInt(rstProjSoum.Fields("CSA")))
1230    chkCUL.Value = Abs(CInt(rstProjSoum.Fields("CUL")))
1235    chkUL.Value = Abs(CInt(rstProjSoum.Fields("UL")))
1240    chkCUR.Value = Abs(CInt(rstProjSoum.Fields("CUR")))
1245    chkUR.Value = Abs(CInt(rstProjSoum.Fields("UR")))
1250    chkCE.Value = Abs(CInt(rstProjSoum.Fields("CE")))

1255    txtPrixTotal.Text = rstProjSoum.Fields("PrixTotal")
1260    txtProfit.Text = rstProjSoum.Fields("total_profit")

1265    If Not IsNull(rstProjSoum.Fields("Delais")) Then
1270      txtDelais.Text = rstProjSoum.Fields("Delais")
1275    Else
1280      txtDelais.Text = "0"
1285    End If

1290    txtCommission.Text = rstProjSoum.Fields("total_commission")

1295    If Not IsNull(rstProjSoum.Fields("MontantForfait")) Then
1300      txtForfait.Text = rstProjSoum.Fields("MontantForfait")

1305      If Not IsNull(rstProjSoum.Fields("InitialeForfait")) Then
1310        lblForfaitInitiale.Caption = rstProjSoum.Fields("InitialeForfait")
1315      Else
1320        lblForfaitInitiale.Caption = ""
1325      End If
1330    Else
1335      txtForfait.Text = ""
1340      lblForfaitInitiale.Caption = ""
1345    End If

1350    Call rstProjSoum.Close
1355    Set rstProjSoum = Nothing
  
        'Affiche les pieces de la soumission
1360    Call RemplirListViewSoumissionProjet(sNoProjet)

1365    If bPrixPieces = True Then
1370      Call UpdatePieces
1375    End If

1380    m_bModeAffichage = False

1385    Call CalculerPrix

1390    Exit Sub

AfficherErreur:

1395    woups "frmProjSoumElec", "RemplirProjSoum", Err, Erl
End Sub

Private Sub RechercherProjSoum(ByVal sNoProjSoum As String)

5       On Error GoTo AfficherErreur

        'Méthode qui recherche une soumission ou un projet dans le combo
        'et qui le sélectionne
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

45      Exit Sub

AfficherErreur:

50      woups "frmProjSoumElec", "RechercherProjSoum", Err, Erl
End Sub

Private Sub RemplirProjSoum()

5       On Error GoTo AfficherErreur

        'Affiche le projet ou la soumission choisie
10      Dim rstProjSoum As ADODB.Recordset
15      Dim rstSoum     As ADODB.Recordset
20      Dim rstClient   As ADODB.Recordset
25      Dim rstContact  As ADODB.Recordset

30      Set rstProjSoum = New ADODB.Recordset
35      Set rstSoum = New ADODB.Recordset
40      Set rstClient = New ADODB.Recordset
45      Set rstContact = New ADODB.Recordset
    
        'Ouvre le recordset selon le type
50      If m_eType = TYPE_PROJET Then
55        Call rstProjSoum.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
    
60        If Not IsNull(rstProjSoum.Fields("IDSoumission")) Then
65          txtNoSoumission.Text = rstProjSoum.Fields("IDSoumission")
70        Else
75          txtNoSoumission.Text = vbNullString
80        End If

85        If Right$(txtNoProjSoum.Text, 2) >= "60" And Right$(txtNoProjSoum.Text, 2) <= "98" Then
90          If Trim(rstProjSoum.Fields("LiaisonChargeable")) <> "" Then
95            m_sLiaison = rstProjSoum.Fields("LiaisonChargeable")
100         Else
105           m_sLiaison = vbNullString

110           Do While Trim$(m_sLiaison) = vbNullString
115             m_sLiaison = InputBox("Quelle est l'extention au projet " & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 3) & " auquel ce projet sera lié?")
120           Loop

125           rstProjSoum.Fields("LiaisonChargeable") = m_sLiaison

130           Call rstProjSoum.Update
135         End If
140       End If
145     Else
150       Call rstProjSoum.Open("SELECT * FROM GRB_SoumissionElec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
155     End If

160     m_bSansTemps = rstProjSoum.Fields("SansTemps")
  
        'Recordset pour avoir le nom du client
165     Call rstClient.Open("SELECT NomClient, IDClient FROM GRB_Client WHERE IDClient = " & rstProjSoum.Fields("IDClient"), g_connData, adOpenDynamic, adLockOptimistic)
  
        'Recordset pour avoir le nom du contact
170     Call rstContact.Open("SELECT NomContact, IDContact FROM GRB_Contact WHERE IDContact = " & rstProjSoum.Fields("IDContact"), g_connData, adOpenDynamic, adLockOptimistic)
      
175     txtClient.Text = rstClient.Fields("NomClient")
  
180     txtcontact.Text = rstContact.Fields("NomContact")
  
185     Call rstClient.Close
190     Set rstClient = Nothing
  
195     Call rstContact.Close
200     Set rstContact = Nothing
         
205     txtProjet.Text = rstProjSoum.Fields("Description")
210     txtNbreManuel.Text = rstProjSoum.Fields("NbreManuel")
215     txtPrixManuel.Text = Conversion(rstProjSoum.Fields("total_manuel"), MODE_PAS_FORMAT)
220     txtTransport.Text = rstProjSoum.Fields("transport")
    
225     txtTotalPieces.Text = Conversion(rstProjSoum.Fields("Total_Piece"), MODE_ARGENT)
230     txtTotalTemps.Text = Conversion(rstProjSoum.Fields("Total_Temps"), MODE_ARGENT)
235     txtPrixTotal.Text = Conversion(rstProjSoum.Fields("PrixTotal"), MODE_ARGENT)
240     txtImprevus.Text = Conversion(rstProjSoum.Fields("Total_Imprevue"), MODE_ARGENT)
245     txtProfit.Text = Conversion(rstProjSoum.Fields("total_profit"), MODE_ARGENT)
250     txtCommission.Text = Conversion(rstProjSoum.Fields("total_commission"), MODE_ARGENT)

255     If Not IsNull(rstProjSoum.Fields("CheminPhotos")) Then
260       txtCheminPhotos.Text = rstProjSoum.Fields("CheminPhotos")
265     Else
270       txtCheminPhotos.Text = vbNullString
275     End If
  
280     chkCSA.Value = Abs(CInt(rstProjSoum.Fields("CSA")))
285     chkCUL.Value = Abs(CInt(rstProjSoum.Fields("CUL")))
290     chkUL.Value = Abs(CInt(rstProjSoum.Fields("UL")))
295     chkCUR.Value = Abs(CInt(rstProjSoum.Fields("CUR")))
300     chkUR.Value = Abs(CInt(rstProjSoum.Fields("UR")))
305     chkCE.Value = Abs(CInt(rstProjSoum.Fields("CE")))
  
310     If Not IsNull(rstProjSoum.Fields("Delais")) Then
315       txtDelais.Text = Trim(rstProjSoum.Fields("Delais"))
320     Else
325       txtDelais.Text = ""
330     End If

335     If Not IsNull(rstProjSoum.Fields("MontantForfait")) Then
340       txtForfait.Text = Conversion(rstProjSoum.Fields("MontantForfait"), MODE_ARGENT)

345       If Not IsNull(rstProjSoum.Fields("InitialeForfait")) Then
350         lblForfaitInitiale.Caption = rstProjSoum.Fields("InitialeForfait")
355       Else
360         lblForfaitInitiale.Caption = ""
365       End If
370     Else
375       txtForfait.Text = ""
380       lblForfaitInitiale.Caption = ""
385     End If

390     If m_eType = TYPE_PROJET Then
395       If Not IsNull(rstProjSoum.Fields("PrixRéception")) Then
400         If Trim(rstProjSoum.Fields("PrixRéception")) <> "" Then
405           txtPrixReception.Text = Conversion(rstProjSoum.Fields("PrixRéception"), MODE_ARGENT)
410         Else
415           txtPrixReception.Text = Conversion("0", MODE_ARGENT)
420         End If
425       Else
430         txtPrixReception.Text = Conversion("0", MODE_ARGENT)
435       End If

440       If Not IsNull(rstProjSoum.Fields("IDSoumission")) Then
445         Call rstSoum.Open("SELECT * FROM GRB_SoumissionElec WHERE IDSoumission = '" & rstProjSoum.Fields("IDSoumission") & "'", g_connData, adOpenDynamic, adLockOptimistic)

450         If Not rstSoum.EOF Then
455           If Not IsNull(rstSoum.Fields("PrixTotal")) Then
460             If Trim(rstSoum.Fields("PrixTotal")) <> "" Then
465               txtPrixSoumission.Text = Conversion(rstSoum.Fields("PrixTotal"), MODE_ARGENT)
470             Else
475               txtPrixSoumission.Text = Conversion(0, MODE_ARGENT)
480             End If
485           Else
490             txtPrixSoumission.Text = Conversion(0, MODE_ARGENT)
495           End If
500         Else
505           txtPrixSoumission.Text = Conversion(0, MODE_ARGENT)
510         End If

515         Call rstSoum.Close
520         Set rstSoum = Nothing
525       End If
530     End If

535     Call rstProjSoum.Close
540     Set rstProjSoum = Nothing
  
        'Affiche les pieces de la soumission
545     Call RemplirListViewProjSoum(txtNoProjSoum.Text)

550     Exit Sub

AfficherErreur:

555     woups "frmProjSoumElec", "RemplirProjSoum", Err, Erl
End Sub

Private Sub RemplirComboCategoriesPieces()

5       On Error GoTo AfficherErreur

        'Remplir le combo des tables (Pièces)
10      Dim rstCatalogueElec As ADODB.Recordset
  
        'Il faut vider le combo avant de le remplir
15      Call cmbPieces.Clear
     
        'Cette méthode crée un recordset contenant les categorie
        'le nom de toutes les tables de la BD
20      Set rstCatalogueElec = New ADODB.Recordset
        
25      Call rstCatalogueElec.Open("SELECT DISTINCT CATEGORIE FROM GRB_CatalogueElec ORDER BY CATEGORIE", g_connData, adOpenDynamic, adLockOptimistic)
    
        'Tant que ce n'est pas la fin des enregistrements
30      Do While Not rstCatalogueElec.EOF
35        Call cmbPieces.AddItem(rstCatalogueElec.Fields("CATEGORIE"))
      
40        Call rstCatalogueElec.MoveNext
45      Loop
    
50      Call rstCatalogueElec.Close
55      Set rstCatalogueElec = Nothing
  
60      If cmbPieces.ListCount > 0 Then
65        cmbPieces.ListIndex = 0
70      End If

75      Exit Sub

AfficherErreur:

80      woups "frmProjSoumElec", "RemplirComboCategoriesPieces", Err, Erl
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

70      woups "frmProjSoumElec", "RemplirComboClients", Err, Erl
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

115     woups "frmProjSoumElec", "RemplirComboContacts", Err, Erl
End Sub

Private Sub RemplirComboSections()

5       On Error GoTo AfficherErreur
        
        'Remplis le combo des sections
10      Dim rstSection As ADODB.Recordset
15      Dim sChamps    As String
  
20      Call cmbSections.Clear
  
25      Set rstSection = New ADODB.Recordset
  
        'Il faut le remplir selon l'ordre
30      Call rstSection.Open("SELECT * FROM GRB_SoumProjSectionElec ORDER BY Ordre", g_connData, adOpenDynamic, adLockOptimistic)
    
35      If m_eLangage = ANGLAIS Then
40        sChamps = "NomSectionEN"
45      Else
50        sChamps = "NomSectionFR"
55      End If
    
60      Do While Not rstSection.EOF
          'On met le nom de la section dans le combo
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

135     woups "frmProjSoumElec", "RemplirComboSections", Err, Erl
End Sub

Private Sub cmdImprimer_Click()

5       On Error GoTo AfficherErreur

10      Dim rstProjSoum As ADODB.Recordset
15      Dim sUser       As String

20      If txtNoProjSoum.Text <> vbNullString Then
25        If VerifierSiOuvert(sUser) = False Then

30          Set rstProjSoum = New ADODB.Recordset
            
            'Ouvre les tables
35          If m_eType = TYPE_PROJET Then
40            If MsgBox("Voulez-vous faire imprimer le projet et tous les extras associés à ce projet?", vbYesNo) = vbYes Then
45              Call rstProjSoum.Open("SELECT * FROM GRB_ProjetElec WHERE Left(IDProjet, 6) = '" & Left$(txtNoProjSoum.Text, 6) & "' ORDER BY IDProjet", g_connData, adOpenDynamic, adLockOptimistic)
50            Else
55              Call rstProjSoum.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & txtNoProjSoum.Text & "' ORDER BY IDProjet", g_connData, adOpenDynamic, adLockOptimistic)
60            End If
65          Else
70            Call rstProjSoum.Open("SELECT * FROM GRB_SoumissionElec WHERE IDSoumission = '" & txtNoProjSoum.Text & "' ORDER BY IDSoumission", g_connData, adOpenDynamic, adLockOptimistic)
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

175     woups "frmProjSoumElec", "cmdImprimer_Click", Err, Erl
End Sub

Private Sub ImprimerProjSoum(ByVal rstProjSoum As ADODB.Recordset)

5       On Error GoTo AfficherErreur

        'Impression de la feuille de soumission
10      Dim rstPiece              As ADODB.Recordset
15      Dim rstPrixSoum           As ADODB.Recordset
20      Dim rstTemp               As ADODB.Recordset
25      Dim rstImpProjSoum        As ADODB.Recordset
30      Dim rstSoum               As ADODB.Recordset
35      Dim sOrdreSection         As String
40      Dim iCompteurSoum         As Integer
45      Dim sSousSection          As String
50      Dim sSousSectionRS        As String
55      Dim dblTempsDessin        As Double
60      Dim dblTempsFabrication   As Double
65      Dim dblTempsAssemblage    As Double
70      Dim dblTempsProgInterface As Double
75      Dim dblTempsProgAutomate  As Double
80      Dim dblTempsProgRobot     As Double
85      Dim dblTempsVision        As Double
90      Dim dblTempsTest          As Double
95      Dim dblTempsInstallation  As Double
100     Dim dblTempsMiseService   As Double
105     Dim dblTempsFormation     As Double
110     Dim dblTempsGestion       As Double
115     Dim dblTempsShipping      As Double
120     Dim dblTotalTemps         As Double
125     Dim dblTotalAutre         As Double
130     Dim dblTotalReste         As Double
135     Dim dblTotalHebergement   As Double
140     Dim dblTotalRepas         As Double
145     Dim dblTotalTransport     As Double
150     Dim dblTotalUniteMobile   As Double
155     Dim sChampsSection        As String
160     Dim sNoProjet             As String
165     Dim sNoSoumission         As String
170     Dim dblPrixEmballage      As Double
      
        'Supprime les données de l'impression
175     Call g_connData.Execute("DELETE * FROM GRB_impression_soumission")
      
180     iCompteurSoum = 1
  
185     Screen.MousePointer = vbHourglass

190     Set rstImpProjSoum = New ADODB.Recordset

195     Call rstImpProjSoum.Open("SELECT * FROM GRB_impression_soumission", g_connData, adOpenDynamic, adLockOptimistic)
    
200     sOrdreSection = vbNullString
            
205     Set rstPiece = New ADODB.Recordset
            
210     If m_eType = TYPE_PROJET Then
215       sNoProjet = rstProjSoum.Fields("IDProjet")

220       If Not IsNull(rstProjSoum.Fields("IDSoumission")) Then
225         sNoSoumission = rstProjSoum.Fields("IDSoumission")
230       Else
235         sNoSoumission = vbNullString
240       End If

245       Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoProjet & "' And Type = 'E' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
250     Else
255       sNoProjet = vbNullString
260       sNoSoumission = rstProjSoum.Fields("IDSoumission")

265       Call rstPiece.Open("SELECT * FROM GRB_Soumission_Pieces WHERE IDSoumission = '" & sNoSoumission & "' And Type = 'E' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
270     End If
  
275     Set rstTemp = New ADODB.Recordset
  
280     Do While Not rstPiece.EOF
285       sSousSectionRS = rstPiece.Fields("SousSection")
       
290       If sSousSectionRS = S_PAS_SOUS_SECTION Then
295         sSousSectionRS = " "
300       End If
      
305       If sOrdreSection <> rstPiece.Fields("OrdreSection") Then
            'remplis la table impression_soumission
            'ajoute seulement la section
310         sOrdreSection = rstPiece.Fields("OrdreSection")
        
315         If m_eLangage = ANGLAIS Then
320           sChampsSection = "NomSectionEN"
325         Else
330           sChampsSection = "NomSectionFR"
335         End If

340         Call rstTemp.Open("SELECT " & sChampsSection & " FROM GRB_SoumProjSectionElec WHERE IDSection = " & rstPiece.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
          
            'Ajoute la section dans la soumission
345         Call rstImpProjSoum.AddNew
          
350         rstImpProjSoum.Fields("NoLigne") = iCompteurSoum

355         If m_eType = TYPE_PROJET Then
360           rstImpProjSoum.Fields("IDSoumission") = sNoProjet
365         Else
370           rstImpProjSoum.Fields("IDSoumission") = sNoSoumission
375         End If
        
380         If Not IsNull(rstTemp.Fields(sChampsSection)) Then
385           rstImpProjSoum.Fields("NomSection") = rstTemp.Fields(sChampsSection)
390         Else
395           rstImpProjSoum.Fields("NomSection") = " "
400         End If
           
405         Call rstImpProjSoum.Update
          
410         iCompteurSoum = iCompteurSoum + 1
        
415         Call rstTemp.Close
          
420         sSousSection = rstPiece.Fields("SousSection")
          
425         If sSousSection = S_PAS_SOUS_SECTION Or sSousSection = "" Then
430           sSousSection = " "
435         End If
          
440         Call rstImpProjSoum.AddNew

445         rstImpProjSoum.Fields("NoLigne") = iCompteurSoum

450         If m_eType = TYPE_PROJET Then
455           rstImpProjSoum.Fields("IDSoumission") = sNoProjet
460         Else
465           rstImpProjSoum.Fields("IDSoumission") = sNoSoumission
470         End If

475         rstImpProjSoum.Fields("SousSection") = sSousSection
        
480         Call rstImpProjSoum.Update
          
485         iCompteurSoum = iCompteurSoum + 1
490       Else
            'ajoute une soussection dans impression_soum
495         If sSousSection <> sSousSectionRS Then
500           sSousSection = sSousSectionRS
      
505           Call rstImpProjSoum.AddNew
        
510           rstImpProjSoum.Fields("NoLigne") = iCompteurSoum

515           If m_eType = TYPE_PROJET Then
520             rstImpProjSoum.Fields("IDSoumission") = sNoProjet
525           Else
530             rstImpProjSoum.Fields("IDSoumission") = sNoSoumission
535           End If

540           rstImpProjSoum.Fields("SousSection") = sSousSectionRS
                    
545           Call rstImpProjSoum.Update
          
550           iCompteurSoum = iCompteurSoum + 1
555         End If
560       End If
          
          'ajoute une piece dans impression_soum
565       Call rstImpProjSoum.AddNew
      
570       rstImpProjSoum.Fields("NoLigne") = iCompteurSoum

575       If m_eType = TYPE_PROJET Then
580         rstImpProjSoum.Fields("IDSoumission") = sNoProjet
585       Else
590         rstImpProjSoum.Fields("IDSoumission") = sNoSoumission
595       End If

600       rstImpProjSoum.Fields("NumItem") = rstPiece.Fields("NumItem")
605       rstImpProjSoum.Fields("Qté") = rstPiece.Fields("Qté")
      
610       If m_eLangage = ANGLAIS Then
615         rstImpProjSoum.Fields("Description") = rstPiece.Fields("DESC_EN")
620       Else
625         rstImpProjSoum.Fields("Description") = rstPiece.Fields("DESC_FR")
630       End If


        '************************************************************************************************
        'SECTION MODIFIER PAR GAÉTAN GINGRAS LE 6 FÉVRIER 2010
        '************************************************************************************************
        
635       rstImpProjSoum.Fields("MANUFACT") = rstPiece.Fields("MANUFACT")
640       'rstImpProjSoum.Fields("PRIX_LIST") = rstPiece.Fields("PRIX_LIST")
            
645       'If Trim(rstPiece.Fields("ESCOMPTE")) <> vbNullString Then
650       '  rstImpProjSoum.Fields("ESCOMPTE") = Conversion(Replace(rstPiece.Fields("ESCOMPTE"), ".", ","), MODE_POURCENT)
655       'Else
660       '  rstImpProjSoum.Fields("ESCOMPTE") = Conversion(0, MODE_POURCENT)
665       'End If
      
670       rstImpProjSoum.Fields("PRIX_NET") = rstPiece.Fields("PRIX_NET")
                     
675       Call rstTemp.Open("SELECT NomFournisseur FROM GRB_Fournisseur WHERE IDFRS = " & rstPiece.Fields("IDFRS"), g_connData, adOpenDynamic, adLockOptimistic)
           
680       If Not rstTemp.EOF Then
685         rstImpProjSoum.Fields("NomFournisseur") = rstTemp.Fields("NomFournisseur")
690       End If
          
695       Call rstTemp.Close

700       'rstImpProjSoum.Fields("TEMPS") = rstPiece.Fields("TEMPS")
705       'rstImpProjSoum.Fields("TEMPS_TOTAL") = rstPiece.Fields("TEMPS_TOTAL")
710       rstImpProjSoum.Fields("PRIX_TOTAL") = rstPiece.Fields("PRIX_TOTAL")
712       rstImpProjSoum.Fields("PROFIT_ARGENT") = rstPiece.Fields("PROFIT_ARGENT")


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
       
713       Call rstImpProjSoum.Update
     
714       iCompteurSoum = iCompteurSoum + 1
    
          'prochaine enreg
715       Call rstPiece.MoveNext
716     Loop
        
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

725     Set DR_SoumissionElec.DataSource = rstImpProjSoum

        '***********************************************************************
        'Pour cas d'urgence, la fonction d'exporter dans Excel va être ici.
        'Nous pourrons plus tard le mettre à un meilleur endroit.
        'Gaétan Gingras le 14 mai 2009
        '***********************************************************************

726     If bTrigger = False Then
727         bTrigger = True
728         intdummie = MsgBox("Désirez-vous exporter les données dans Excel, SEULEMENT ?", vbYesNo + vbInformation, "Exportation dans Excel")
729        End If
730     If intdummie = vbYes Then
742         Dim sqlstr As String
743         Dim rstExport As ADODB.Recordset
744         Set rstExport = New ADODB.Recordset
745         'sqlstr = "SELECT GRB_impression_soumission.IDSoumission, CDbl([Qté]) AS Quantité, GRB_impression_soumission.NumItem, GRB_impression_soumission.Description, GRB_impression_soumission.Manufact, CDbl([Prix_list]) AS PrixdeListe, CDbl(Left([escompte],Len([escompte])-1)) AS Escomptes, CDbl([Prix_net]) AS prix_nette, GRB_impression_soumission.NomFournisseur, GRB_impression_soumission.DateReception , GRB_impression_soumission.DateCommande "
746         sqlstr = "SELECT GRB_impression_soumission.IDSoumission, CDbl([Qté]) AS Quantité, GRB_impression_soumission.NumItem, GRB_impression_soumission.Description, GRB_impression_soumission.Manufact, CDbl([Prix_list]) AS PrixdeListe, CDbl(Left([escompte],Len([escompte])-1)) AS Escomptes, CDbl([Prix_net]) AS prix_nette, GRB_impression_soumission.Prix_total - GRB_impression_soumission.Profit_Argent AS Prix_Total ,GRB_impression_soumission.NomFournisseur, GRB_impression_soumission.DateReception , GRB_impression_soumission.DateCommande ,  GRB_impression_soumission.NoSéquentiel "

            sqlstr = sqlstr + "FROM GRB_impression_soumission "
747         sqlstr = sqlstr + "WHERE (((GRB_impression_soumission.IDSoumission)='" & sProjet & "') AND ((GRB_impression_soumission.NumItem) Is Not Null)) "
748         sqlstr = sqlstr + "ORDER BY GRB_impression_soumission.noligne"
749         Call rstExport.Open(sqlstr, g_connData, adOpenDynamic, adLockOptimistic)
750         Call ExportdansExcel(rstExport)
751         Screen.MousePointer = vbDefault
752         Exit Sub
753      End If
        '***********************************************************************
      
775     Call TraduireImpressionSoumission

780     If m_eType = TYPE_PROJET Then
785       DR_SoumissionElec.Sections("Section5").Controls("shpCadrePrixReception").Visible = True
790       DR_SoumissionElec.Sections("Section5").Controls("lblTitrePrixReception").Visible = True
795       DR_SoumissionElec.Sections("Section5").Controls("lblPrixReception").Visible = True

800       DR_SoumissionElec.Sections("Section5").Controls("shpCadrePrixSoumission").Visible = True
805       DR_SoumissionElec.Sections("Section5").Controls("lblTitrePrixSoumission").Visible = True
810       DR_SoumissionElec.Sections("Section5").Controls("lblPrixSoumission").Visible = True
815     Else
820       DR_SoumissionElec.Sections("Section5").Controls("shpCadrePrixReception").Visible = False
825       DR_SoumissionElec.Sections("Section5").Controls("lblTitrePrixReception").Visible = False
830       DR_SoumissionElec.Sections("Section5").Controls("lblPrixReception").Visible = False

835       DR_SoumissionElec.Sections("Section5").Controls("shpCadrePrixSoumission").Visible = False
840       DR_SoumissionElec.Sections("Section5").Controls("lblTitrePrixSoumission").Visible = False
845       DR_SoumissionElec.Sections("Section5").Controls("lblPrixSoumission").Visible = False
850     End If
                              
        'affiche la date
        '**************************************************
        'ajout par Gaétan Gingras le 20 mai 2009
854     If MsgBox("Désirez-vous afficher la date en bas de page ?", vbYesNo + vbInformation, "Affichage de la date") = vbYes Then
855         DR_SoumissionElec.Sections("section3").Controls("lbldate").Caption = ConvertDate(Date)
856     Else
857         DR_SoumissionElec.Sections("section3").Controls("lbldate").Caption = " "
858     End If
        '**************************************************
        
        'affiche entete
860     If Not IsNull(rstProjSoum.Fields("IDSoumission")) Then
865       DR_SoumissionElec.Sections("Section2").Controls("lblSoumission").Caption = rstProjSoum.Fields("IDSoumission")
870     Else
875       DR_SoumissionElec.Sections("Section2").Controls("lblSoumission").Caption = vbNullString
880     End If
                  
885     If m_eType = TYPE_PROJET Then
890       DR_SoumissionElec.Sections("Section2").Controls("lblprojet").Caption = rstProjSoum.Fields("IDProjet")
895     Else
900       DR_SoumissionElec.Sections("Section2").Controls("lblProjet").Caption = vbNullString
905     End If
                
910     DR_SoumissionElec.Sections("Section2").Controls("lbldescription").Caption = rstProjSoum.Fields("Description")
    
915     Call rstTemp.Open("SELECT NomClient FROM GRB_Client WHERE IDClient = " & rstProjSoum.Fields("IDClient"), g_connData, adOpenDynamic, adLockOptimistic)
      
920     DR_SoumissionElec.Sections("Section2").Controls("lblclient").Caption = rstTemp.Fields("NomClient")

925     Call rstTemp.Close
      
930     Call rstTemp.Open("SELECT NomContact FROM GRB_Contact WHERE IDContact = " & rstProjSoum.Fields("IDContact"), g_connData, adOpenDynamic, adLockOptimistic)
      
935     DR_SoumissionElec.Sections("Section2").Controls("lblcontact").Caption = rstTemp.Fields("NomContact")
                   
940     Call rstTemp.Close
      
        'Affiche pied d'état
     
        'Temps
945     If Not IsNull(rstProjSoum.Fields("TauxDessin")) Then
950       DR_SoumissionElec.Sections("Section5").Controls("lblTauxDessin").Caption = rstProjSoum.Fields("TauxDessin")
955     Else
960       DR_SoumissionElec.Sections("Section5").Controls("lblTauxDessin").Caption = "0"
965     End If

970     If Not IsNull(rstProjSoum.Fields("TauxFabrication")) Then
975       DR_SoumissionElec.Sections("Section5").Controls("lblTauxFabrication").Caption = rstProjSoum.Fields("TauxFabrication")
980     Else
985       DR_SoumissionElec.Sections("Section5").Controls("lblTauxFabrication").Caption = "0"
990     End If

995     If Not IsNull(rstProjSoum.Fields("TauxAssemblage")) Then
1000      DR_SoumissionElec.Sections("Section5").Controls("lblTauxAssemblage").Caption = rstProjSoum.Fields("TauxAssemblage")
1005    Else
1010      DR_SoumissionElec.Sections("Section5").Controls("lblTauxAssemblage").Caption = "0"
1015    End If

1020    If Not IsNull(rstProjSoum.Fields("TauxProgInterface")) Then
1025      DR_SoumissionElec.Sections("Section5").Controls("lblTauxProgInterface").Caption = rstProjSoum.Fields("TauxProgInterface")
1030    Else
1035      DR_SoumissionElec.Sections("Section5").Controls("lblTauxProgInterface").Caption = "0"
1040    End If

1045    If Not IsNull(rstProjSoum.Fields("TauxProgAutomate")) Then
1050      DR_SoumissionElec.Sections("Section5").Controls("lblTauxProgAutomate").Caption = rstProjSoum.Fields("TauxProgAutomate")
1055    Else
1060      DR_SoumissionElec.Sections("Section5").Controls("lblTauxProgAutomate").Caption = "0"
1065    End If

1070    If Not IsNull(rstProjSoum.Fields("TauxProgRobot")) Then
1075      DR_SoumissionElec.Sections("Section5").Controls("lblTauxProgRobot").Caption = rstProjSoum.Fields("TauxProgRobot")
1080    Else
1085      DR_SoumissionElec.Sections("Section5").Controls("lblTauxProgRobot").Caption = "0"
1090    End If

1095    If Not IsNull(rstProjSoum.Fields("TauxVision")) Then
1100      DR_SoumissionElec.Sections("Section5").Controls("lblTauxVision").Caption = rstProjSoum.Fields("TauxVision")
1105    Else
1110      DR_SoumissionElec.Sections("Section5").Controls("lblTauxVision").Caption = "0"
1115    End If

1120    If Not IsNull(rstProjSoum.Fields("TauxTest")) Then
1125      DR_SoumissionElec.Sections("Section5").Controls("lblTauxTest").Caption = rstProjSoum.Fields("TauxTest")
1130    Else
1135      DR_SoumissionElec.Sections("Section5").Controls("lblTauxTest").Caption = "0"
1140    End If

1145    If Not IsNull(rstProjSoum.Fields("TauxInstallation")) Then
1150      DR_SoumissionElec.Sections("Section5").Controls("lblTauxInstallation").Caption = rstProjSoum.Fields("TauxInstallation")
1155    Else
1160      DR_SoumissionElec.Sections("Section5").Controls("lblTauxInstallation").Caption = "0"
1165    End If

1170    If Not IsNull(rstProjSoum.Fields("TauxMiseService")) Then
1175      DR_SoumissionElec.Sections("Section5").Controls("lblTauxMiseService").Caption = rstProjSoum.Fields("TauxMiseService")
1180    Else
1185      DR_SoumissionElec.Sections("Section5").Controls("lblTauxMiseService").Caption = "0"
1190    End If

1195    If Not IsNull(rstProjSoum.Fields("TauxFormation")) Then
1200      DR_SoumissionElec.Sections("Section5").Controls("lblTauxFormation").Caption = rstProjSoum.Fields("TauxFormation")
1205    Else
1210      DR_SoumissionElec.Sections("Section5").Controls("lblTauxFormation").Caption = "0"
1215    End If

1220    If Not IsNull(rstProjSoum.Fields("TauxGestion")) Then
1225      DR_SoumissionElec.Sections("Section5").Controls("lblTauxGestion").Caption = rstProjSoum.Fields("TauxGestion")
1230    Else
1235      DR_SoumissionElec.Sections("Section5").Controls("lblTauxGestion").Caption = "0"
1240    End If

1245    If Not IsNull(rstProjSoum.Fields("TauxShipping")) Then
1250      DR_SoumissionElec.Sections("Section5").Controls("lblTauxShipping").Caption = rstProjSoum.Fields("TauxShipping")
1255    Else
1260      DR_SoumissionElec.Sections("Section5").Controls("lblTauxShipping").Caption = "0"
1265    End If

1270    If Not IsNull(rstProjSoum.Fields("TempsDessin")) Then
1275      DR_SoumissionElec.Sections("Section5").Controls("lblTempsDessinSoum").Caption = rstProjSoum.Fields("TempsDessin")
1280    Else
1285      DR_SoumissionElec.Sections("Section5").Controls("lblTempsDessinSoum").Caption = "0"
1290    End If

1295    If Not IsNull(rstProjSoum.Fields("TempsFabrication")) Then
1300      If rstProjSoum.Fields("SansTemps") = False Then
1305        DR_SoumissionElec.Sections("Section5").Controls("lblTempsFabricationSoum").Caption = rstProjSoum.Fields("TempsFabrication")
1310      Else
1315        DR_SoumissionElec.Sections("Section5").Controls("lblTempsFabricationSoum").Caption = "0"
1320      End If
1325    Else
1330      DR_SoumissionElec.Sections("Section5").Controls("lblTempsFabricationSoum").Caption = "0"
1335    End If

1340    If Not IsNull(rstProjSoum.Fields("TempsAssemblage")) Then
1345      DR_SoumissionElec.Sections("Section5").Controls("lblTempsAssemblageSoum").Caption = rstProjSoum.Fields("TempsAssemblage")
1350    Else
1355      DR_SoumissionElec.Sections("Section5").Controls("lblTempsAssemblageSoum").Caption = "0"
1360    End If

1365    If Not IsNull(rstProjSoum.Fields("TempsProgInterface")) Then
1370      DR_SoumissionElec.Sections("Section5").Controls("lblTempsProgInterfaceSoum").Caption = rstProjSoum.Fields("TempsProgInterface")
1375    Else
1380      DR_SoumissionElec.Sections("Section5").Controls("lblTempsProgInterfaceSoum").Caption = "0"
1385    End If

1390    If Not IsNull(rstProjSoum.Fields("TempsProgAutomate")) Then
1395      DR_SoumissionElec.Sections("Section5").Controls("lblTempsProgAutomateSoum").Caption = rstProjSoum.Fields("TempsProgAutomate")
1400    Else
1405      DR_SoumissionElec.Sections("Section5").Controls("lblTempsProgAutomateSoum").Caption = "0"
1410    End If

1415    If Not IsNull(rstProjSoum.Fields("TempsProgRobot")) Then
1420      DR_SoumissionElec.Sections("Section5").Controls("lblTempsProgRobotSoum").Caption = rstProjSoum.Fields("TempsProgRobot")
1425    Else
1430      DR_SoumissionElec.Sections("Section5").Controls("lblTempsProgRobotSoum").Caption = "0"
1435    End If

1440    If Not IsNull(rstProjSoum.Fields("TempsVision")) Then
1445      DR_SoumissionElec.Sections("Section5").Controls("lblTempsVisionSoum").Caption = rstProjSoum.Fields("TempsVision")
1450    Else
1455      DR_SoumissionElec.Sections("Section5").Controls("lblTempsVisionSoum").Caption = "0"
1460    End If

1465    If Not IsNull(rstProjSoum.Fields("TempsTest")) Then
1470      DR_SoumissionElec.Sections("Section5").Controls("lblTempsTestSoum").Caption = rstProjSoum.Fields("TempsTest")
1475    Else
1480      DR_SoumissionElec.Sections("Section5").Controls("lblTempsTestSoum").Caption = "0"
1485    End If

1490    If Not IsNull(rstProjSoum.Fields("TempsInstallation")) Then
1495      DR_SoumissionElec.Sections("Section5").Controls("lblTempsInstallationSoum").Caption = rstProjSoum.Fields("TempsInstallation")
1500    Else
1505      DR_SoumissionElec.Sections("Section5").Controls("lblTempsInstallationSoum").Caption = "0"
1510    End If

1515    If Not IsNull(rstProjSoum.Fields("TempsMiseService")) Then
1520      DR_SoumissionElec.Sections("Section5").Controls("lblTempsMiseServiceSoum").Caption = rstProjSoum.Fields("TempsMiseService")
1525    Else
1530      DR_SoumissionElec.Sections("Section5").Controls("lblTempsMiseServiceSoum").Caption = "0"
1535    End If

1540    If Not IsNull(rstProjSoum.Fields("TempsFormation")) Then
1545      DR_SoumissionElec.Sections("Section5").Controls("lblTempsFormationSoum").Caption = rstProjSoum.Fields("TempsFormation")
1550    Else
1555      DR_SoumissionElec.Sections("Section5").Controls("lblTempsFormationSoum").Caption = "0"
1560    End If

1565    If Not IsNull(rstProjSoum.Fields("TempsGestion")) Then
1570      DR_SoumissionElec.Sections("Section5").Controls("lblTempsGestionSoum").Caption = rstProjSoum.Fields("TempsGestion")
1575    Else
1580      DR_SoumissionElec.Sections("Section5").Controls("lblTempsGestionSoum").Caption = "0"
1585    End If

1590    If Not IsNull(rstProjSoum.Fields("TempsShipping")) Then
1595      DR_SoumissionElec.Sections("Section5").Controls("lblTempsShippingSoum").Caption = rstProjSoum.Fields("TempsShipping")
1600    Else
1605      DR_SoumissionElec.Sections("Section5").Controls("lblTempsShippingSoum").Caption = "0"
1610    End If

1615    If Not IsNull(rstProjSoum.Fields("TempsDessin")) Then
1620      If IsNumeric(rstProjSoum.Fields("TempsDessin")) Then
1625        dblTempsDessin = CDbl(rstProjSoum.Fields("TempsDessin"))
1630      Else
1635        dblTempsDessin = 0
1640      End If
1645    Else
1650      dblTempsDessin = 0
1655    End If

1660    If Not IsNull(rstProjSoum.Fields("TempsFabrication")) Then
1665      If rstProjSoum.Fields("SansTemps") = False Then
1670        If IsNumeric(rstProjSoum.Fields("TempsFabrication")) Then
1675          dblTempsFabrication = CDbl(rstProjSoum.Fields("TempsFabrication"))
1680        Else
1685          dblTempsFabrication = 0
1690        End If
1695      Else
1700        dblTempsFabrication = 0
1705      End If
1710    Else
1715      dblTempsFabrication = 0
1720    End If

1725    If Not IsNull(rstProjSoum.Fields("TempsAssemblage")) Then
1730      If IsNumeric(rstProjSoum.Fields("TempsAssemblage")) Then
1735        dblTempsAssemblage = CDbl(rstProjSoum.Fields("TempsAssemblage"))
1740      Else
1745        dblTempsAssemblage = 0
1750      End If
1755    Else
1760      dblTempsAssemblage = 0
1765    End If

1770    If Not IsNull(rstProjSoum.Fields("TempsProgInterface")) Then
1775      If IsNumeric(rstProjSoum.Fields("TempsProgInterface")) Then
1780        dblTempsProgInterface = CDbl(rstProjSoum.Fields("TempsProgInterface"))
1785      Else
1790        dblTempsProgInterface = 0
1795      End If
1800    Else
1805      dblTempsProgInterface = 0
1810    End If

1815    If Not IsNull(rstProjSoum.Fields("TempsProgAutomate")) Then
1820      If IsNumeric(rstProjSoum.Fields("TempsProgAutomate")) Then
1825        dblTempsProgAutomate = CDbl(rstProjSoum.Fields("TempsProgAutomate"))
1830      Else
1835        dblTempsProgAutomate = 0
1840      End If
1845    Else
1850      dblTempsProgAutomate = 0
1855    End If

1860    If Not IsNull(rstProjSoum.Fields("TempsProgRobot")) Then
1865      If IsNumeric(rstProjSoum.Fields("TempsProgRobot")) Then
1870        dblTempsProgRobot = CDbl(rstProjSoum.Fields("TempsProgRobot"))
1875      Else
1880        dblTempsProgRobot = 0
1885      End If
1890    Else
1895      dblTempsProgRobot = 0
1900    End If

1905    If Not IsNull(rstProjSoum.Fields("TempsVision")) Then
1910      If IsNumeric(rstProjSoum.Fields("TempsVision")) Then
1915        dblTempsVision = CDbl(rstProjSoum.Fields("TempsVision"))
1920      Else
1925        dblTempsVision = 0
1930      End If
1935    Else
1940      dblTempsVision = 0
1945    End If

1950    If Not IsNull(rstProjSoum.Fields("TempsTest")) Then
1955      If IsNumeric(rstProjSoum.Fields("TempsTest")) Then
1960        dblTempsTest = CDbl(rstProjSoum.Fields("TempsTest"))
1965      Else
1970        dblTempsTest = 0
1975      End If
1980    Else
1985      dblTempsTest = 0
1990    End If

1995    If Not IsNull(rstProjSoum.Fields("TempsInstallation")) Then
2000      If IsNumeric(rstProjSoum.Fields("TempsInstallation")) Then
2005        dblTempsInstallation = CDbl(rstProjSoum.Fields("TempsInstallation"))
2010      Else
2015        dblTempsInstallation = 0
2020      End If
2025    Else
2030      dblTempsInstallation = 0
2035    End If

2040    If Not IsNull(rstProjSoum.Fields("TempsMiseService")) Then
2045      If IsNumeric(rstProjSoum.Fields("TempsMiseService")) Then
2050        dblTempsMiseService = CDbl(rstProjSoum.Fields("TempsMiseService"))
2055      Else
2060        dblTempsMiseService = 0
2065      End If
2070    Else
2075      dblTempsMiseService = 0
2080    End If

2085    If Not IsNull(rstProjSoum.Fields("TempsFormation")) Then
2090      If IsNumeric(rstProjSoum.Fields("TempsFormation")) Then
2095        dblTempsFormation = CDbl(rstProjSoum.Fields("TempsFormation"))
2100      Else
2105        dblTempsFormation = 0
2110      End If
2115    Else
2120      dblTempsFormation = 0
2125    End If

2130    If Not IsNull(rstProjSoum.Fields("TempsGestion")) Then
2135      If IsNumeric(rstProjSoum.Fields("TempsGestion")) Then
2140        dblTempsGestion = CDbl(rstProjSoum.Fields("TempsGestion"))
2145      Else
2150        dblTempsGestion = 0
2155      End If
2160    Else
2165      dblTempsGestion = 0
2170    End If

2175    If Not IsNull(rstProjSoum.Fields("TempsShipping")) Then
2180      If IsNumeric(rstProjSoum.Fields("TempsShipping")) Then
2185        dblTempsShipping = CDbl(rstProjSoum.Fields("TempsShipping"))
2190      Else
2195        dblTempsShipping = 0
2200      End If
2205    Else
2210      dblTempsShipping = 0
2215    End If

2220    dblTotalTemps = dblTempsDessin + _
                        dblTempsFabrication + _
                        dblTempsAssemblage + _
                        dblTempsProgInterface + _
                        dblTempsProgAutomate + _
                        dblTempsProgRobot + _
                        dblTempsVision + _
                        dblTempsTest + _
                        dblTempsInstallation + _
                        dblTempsMiseService + _
                        dblTempsFormation + _
                        dblTempsGestion + _
                        dblTempsShipping
                          
2225    DR_SoumissionElec.Sections("Section5").Controls("lblTotalTempsRHSoum").Caption = dblTotalTemps

2230    If m_eType = TYPE_PROJET Then
2235      Call CalculerTempsReelsImpression(rstProjSoum.Fields("IDProjet"))
2240    End If
    
        'Autres frais
2245    If m_eType = TYPE_PROJET Then
2250      DR_SoumissionElec.Sections("Section5").Controls("lblNbrePersonne").Caption = "0"
2255      DR_SoumissionElec.Sections("Section5").Controls("lblTempsHebergement").Caption = "0"
2260      DR_SoumissionElec.Sections("Section5").Controls("lblTauxHebergement1").Caption = "0"
2265      DR_SoumissionElec.Sections("Section5").Controls("lblTauxHebergement2").Caption = "0"
2270      DR_SoumissionElec.Sections("Section5").Controls("lblTempsRepas").Caption = "0"
2275      DR_SoumissionElec.Sections("Section5").Controls("lblTauxRepas").Caption = "0"
2280      DR_SoumissionElec.Sections("Section5").Controls("lblTempsTransport").Caption = "0"
2285      DR_SoumissionElec.Sections("Section5").Controls("lblTauxTransport").Caption = "0"
2290      DR_SoumissionElec.Sections("Section5").Controls("lblTempsUniteMobile").Caption = "0"
2295      DR_SoumissionElec.Sections("Section5").Controls("lblTauxUniteMobile").Caption = "0"
2300    Else
2305      If Not IsNull(rstProjSoum.Fields("NbrePersonne")) Then
2310        DR_SoumissionElec.Sections("Section5").Controls("lblNbrePersonne").Caption = rstProjSoum.Fields("NbrePersonne")
2315      Else
2320        DR_SoumissionElec.Sections("Section5").Controls("lblNbrePersonne").Caption = "0"
2325      End If

2330      If Not IsNull(rstProjSoum.Fields("TempsHebergement")) Then
2335        DR_SoumissionElec.Sections("Section5").Controls("lblTempsHebergement").Caption = rstProjSoum.Fields("TempsHebergement")
2340      Else
2345        DR_SoumissionElec.Sections("Section5").Controls("lblTempsHebergement").Caption = "0"
2350      End If

2355      If Not IsNull(rstProjSoum.Fields("TauxHebergement1")) Then
2360        DR_SoumissionElec.Sections("Section5").Controls("lblTauxHebergement1").Caption = rstProjSoum.Fields("TauxHebergement1")
2365      Else
2370        DR_SoumissionElec.Sections("Section5").Controls("lblTauxHebergement1").Caption = "0"
2375      End If

2380      If Not IsNull(rstProjSoum.Fields("TauxHebergement2")) Then
2385        DR_SoumissionElec.Sections("Section5").Controls("lblTauxHebergement2").Caption = rstProjSoum.Fields("TauxHebergement2")
2390      Else
2395        DR_SoumissionElec.Sections("Section5").Controls("lblTauxHebergement2").Caption = "0"
2400      End If
           
2405      If Not IsNull(rstProjSoum.Fields("TempsRepas")) Then
2410        DR_SoumissionElec.Sections("Section5").Controls("lblTempsRepas").Caption = rstProjSoum.Fields("TempsRepas")
2415      Else
2420        DR_SoumissionElec.Sections("Section5").Controls("lblTempsRepas").Caption = "0"
2425      End If

2430      If Not IsNull(rstProjSoum.Fields("TauxRepas")) Then
2435        DR_SoumissionElec.Sections("Section5").Controls("lblTauxRepas").Caption = rstProjSoum.Fields("TauxRepas")
2440      Else
2445        DR_SoumissionElec.Sections("Section5").Controls("lblTauxRepas").Caption = "0"
2450      End If

2455      If Not IsNull(rstProjSoum.Fields("TempsTransport")) Then
2460        DR_SoumissionElec.Sections("Section5").Controls("lblTempsTransport").Caption = rstProjSoum.Fields("TempsTransport")
2465      Else
2470        DR_SoumissionElec.Sections("Section5").Controls("lblTempsTransport").Caption = "0"
2475      End If

2480      If Not IsNull(rstProjSoum.Fields("TauxTransport")) Then
2485        DR_SoumissionElec.Sections("Section5").Controls("lblTauxTransport").Caption = rstProjSoum.Fields("TauxTransport")
2490      Else
2495        DR_SoumissionElec.Sections("Section5").Controls("lblTauxTransport").Caption = "0"
2500      End If

2505      If Not IsNull(rstProjSoum.Fields("TempsUniteMobile")) Then
2510        DR_SoumissionElec.Sections("Section5").Controls("lblTempsUniteMobile").Caption = rstProjSoum.Fields("TempsUniteMobile")
2515      Else
2520        DR_SoumissionElec.Sections("Section5").Controls("lblTempsUniteMobile").Caption = "0"
2525      End If

2530      If Not IsNull(rstProjSoum.Fields("TauxUniteMobile")) Then
2535        DR_SoumissionElec.Sections("Section5").Controls("lblTauxUniteMobile").Caption = rstProjSoum.Fields("TauxUniteMobile")
2540      Else
2545        DR_SoumissionElec.Sections("Section5").Controls("lblTauxUniteMobile").Caption = "0"
2550      End If
2555    End If

2560    If Not IsNull(rstProjSoum.Fields("PrixEmballage")) Then
2565      DR_SoumissionElec.Sections("Section5").Controls("lblPrixEmballage").Caption = rstProjSoum.Fields("PrixEmballage")
2570    Else
2575      DR_SoumissionElec.Sections("Section5").Controls("lblPrixEmballage").Caption = "0"
2580    End If

2585    DR_SoumissionElec.Sections("Section5").Controls("lblPrixManuel").Caption = rstProjSoum.Fields("Total_manuel")

2590    DR_SoumissionElec.Sections("Section5").Controls("lblTotalTempsAR").Caption = Conversion(rstProjSoum.Fields("total_temps"), MODE_ARGENT)
2595    DR_SoumissionElec.Sections("Section5").Controls("lblTotalPieceAR").Caption = Conversion(rstProjSoum.Fields("total_piece"), MODE_ARGENT)
2600    DR_SoumissionElec.Sections("Section5").Controls("lblProfit").Caption = Conversion((rstProjSoum.Fields("profit") - 1) * 100, MODE_POURCENT)
2605    DR_SoumissionElec.Sections("Section5").Controls("lblTotalProfit").Caption = Conversion(rstProjSoum.Fields("total_profit"), MODE_ARGENT)
2610    DR_SoumissionElec.Sections("Section5").Controls("lblImprevue").Caption = Conversion(rstProjSoum.Fields("imprevue"), MODE_POURCENT)
2615    DR_SoumissionElec.Sections("Section5").Controls("lblImprevueAR").Caption = Conversion(rstProjSoum.Fields("total_imprevue"), MODE_ARGENT)
2620    DR_SoumissionElec.Sections("Section5").Controls("lblCommission").Caption = Conversion(rstProjSoum.Fields("commission"), MODE_POURCENT)
2625    DR_SoumissionElec.Sections("Section5").Controls("lblCommissionAR").Caption = Conversion(rstProjSoum.Fields("total_commission"), MODE_ARGENT)

2630    If m_eType = TYPE_PROJET Then
2635      If Not IsNull(rstProjSoum.Fields("PrixRéception")) Then
2640        DR_SoumissionElec.Sections("Section5").Controls("lblPrixReception").Caption = Conversion(rstProjSoum.Fields("PrixRéception"), MODE_ARGENT)
2645      Else
2650        DR_SoumissionElec.Sections("Section5").Controls("lblPrixReception").Caption = Conversion(0, MODE_ARGENT)
2655      End If

2660      If Not IsNull(rstProjSoum.Fields("IDSoumission")) Then
2665        Set rstPrixSoum = New ADODB.Recordset

2670        Call rstPrixSoum.Open("SELECT PrixTotal FROM GRB_SoumissionElec WHERE IDSoumission = '" & rstProjSoum.Fields("IDSoumission") & "'", g_connData, adOpenDynamic, adLockOptimistic)

2675        If Not rstPrixSoum.EOF Then
2680          If Not IsNull(rstPrixSoum.Fields("PrixTotal")) Then
2685            DR_SoumissionElec.Sections("Section5").Controls("lblPrixSoumission").Caption = Conversion(rstPrixSoum.Fields("PrixTotal"), MODE_ARGENT)
2690          Else
2695            DR_SoumissionElec.Sections("Section5").Controls("lblPrixSoumission").Caption = Conversion("0", MODE_ARGENT)
2700          End If
2705        Else
2710          DR_SoumissionElec.Sections("Section5").Controls("lblPrixSoumission").Caption = Conversion("0", MODE_ARGENT)
2715        End If

2720        Call rstPrixSoum.Close
2725        Set rstPrixSoum = Nothing
2730      Else
2735        DR_SoumissionElec.Sections("Section5").Controls("lblPrixSoumission").Caption = Conversion("0", MODE_ARGENT)
2740      End If
2745    End If
  
2750    If m_eType = TYPE_PROJET Then
2755      dblTotalHebergement = 0
2760      dblTotalRepas = 0
2765      dblTotalTransport = 0
2770      dblTotalUniteMobile = 0
2775    Else
2780      If Not IsNull(rstProjSoum.Fields("TotalHebergement")) Then
2785        dblTotalHebergement = rstProjSoum.Fields("TotalHebergement")
2790      End If
  
2795      If Not IsNull(rstProjSoum.Fields("TotalRepas")) Then
2800        dblTotalRepas = rstProjSoum.Fields("TotalRepas")
2805      End If
  
2810      If Not IsNull(rstProjSoum.Fields("TempsTransport")) And Not IsNull(rstProjSoum.Fields("TauxTransport")) Then
2815        dblTotalTransport = CDbl(rstProjSoum.Fields("TempsTransport")) * CDbl(rstProjSoum.Fields("TauxTransport"))
2820      Else
2825        dblTotalTransport = 0
2830      End If

2835      If Not IsNull(rstProjSoum.Fields("TempsUniteMobile")) And Not IsNull(rstProjSoum.Fields("TauxUniteMobile")) Then
2840        dblTotalUniteMobile = CDbl(rstProjSoum.Fields("TempsUniteMobile")) * CDbl(rstProjSoum.Fields("TauxUniteMobile"))
2845      Else
2850        dblTotalUniteMobile = 0
2855      End If
2860    End If

2865    If Not IsNull(rstProjSoum.Fields("PrixEmballage")) Then
2870      dblPrixEmballage = CDbl(Replace(rstProjSoum.Fields("PrixEmballage"), ".", ","))
2875    Else
2880      dblPrixEmballage = 0
2885    End If
        
2890    dblTotalReste = dblTotalHebergement + dblTotalRepas + dblTotalTransport + dblTotalUniteMobile + dblPrixEmballage

2895    dblTotalAutre = dblTotalReste + CDbl(rstProjSoum.Fields("total_manuel"))
    
2900    DR_SoumissionElec.Sections("Section5").Controls("lblAutre").Caption = Conversion(CStr(dblTotalAutre), MODE_ARGENT)
    
2905    DR_SoumissionElec.Sections("Section5").Controls("lblGrandTotalAR").Caption = Conversion(rstProjSoum.Fields("prixtotal"), MODE_ARGENT)
            
2910    If rstProjSoum.Fields("MontantForfait") <> "" Then
2915      DR_SoumissionElec.Sections("Section5").Controls("shpCadreForfait").Visible = True
2920      DR_SoumissionElec.Sections("Section5").Controls("lblTitreForfait").Visible = True
2925      DR_SoumissionElec.Sections("Section5").Controls("lblForfait").Visible = True

2930      DR_SoumissionElec.Sections("Section5").Controls("lblTitreForfait").Caption = DR_SoumissionElec.Sections("Section5").Controls("lblTitreForfait").Caption & " ( " & rstProjSoum.Fields("InitialeForfait") & " )"
2935      DR_SoumissionElec.Sections("Section5").Controls("lblForfait").Caption = rstProjSoum.Fields("MontantForfait")
2940    Else
2945      DR_SoumissionElec.Sections("Section5").Controls("shpCadreForfait").Visible = False
2950      DR_SoumissionElec.Sections("Section5").Controls("lblTitreForfait").Visible = False
2955      DR_SoumissionElec.Sections("Section5").Controls("lblForfait").Visible = False
2960    End If

        '************************************************************************************************
        'SECTION MODIFIÉ PAR GAÉTAN GINGRAS LE 7 FÉVRIER 2010
        '************************************************************************************************
        If bFlag = True Then
            DR_SoumissionElec.Sections("Section2").Controls("lbl_DateCommande").Visible = True
            DR_SoumissionElec.Sections("Section2").Controls("lbl_DateReception").Visible = True
            DR_SoumissionElec.Sections("Section1").Controls("txt_DateCommande").Visible = True
            DR_SoumissionElec.Sections("Section1").Controls("txt_DateReception").Visible = True
        Else
            DR_SoumissionElec.Sections("Section2").Controls("lbl_DateCommande").Visible = False
            DR_SoumissionElec.Sections("Section2").Controls("lbl_DateReception").Visible = False
            DR_SoumissionElec.Sections("Section1").Controls("txt_DateCommande").Visible = False
            DR_SoumissionElec.Sections("Section1").Controls("txt_DateReception").Visible = False
        End If
        '************************************************************************************************
        'FIN DE LA SECTION MODIFIÉ
        '************************************************************************************************
        
            
        'Affiche le rapport soumission
2965    DR_SoumissionElec.Orientation = rptOrientLandscape
    
2970    Call DR_SoumissionElec.Show(vbModal)
             
2975    Call rstImpProjSoum.Close
2980    Set rstImpProjSoum = Nothing

2985    Set rstTemp = Nothing
    
2990    Screen.MousePointer = vbDefault

2995    Exit Sub

AfficherErreur:

3000    woups "frmProjSoumElec", "ImprimerProjSoum", Err, Erl
End Sub

Private Sub ImprimerListePieces(ByVal rstProjSoum As ADODB.Recordset)

5       On Error GoTo AfficherErreur

        'Impression de la feuille de la liste des pièces
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

65      Set rstPiece = New ADODB.Recordset
70      Set rstTemp = New ADODB.Recordset
75      Set rstImpListePiece = New ADODB.Recordset

80      Call g_connData.Execute("DELETE * FROM GRB_impression_listepiece")

85      iCompteurPiece = 1

90      Screen.MousePointer = vbHourglass

        'Ouverture du recordset
95      If m_eType = TYPE_PROJET Then
100       sNoProjet = rstProjSoum.Fields("IDProjet")

105       If Not IsNull(rstProjSoum.Fields("IDSoumission")) Then
110         sNoSoumission = rstProjSoum.Fields("IDSoumission")
115       Else
120         sNoSoumission = vbNullString
125       End If

130       Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoProjet & "' AND Type = 'E' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
135     Else
140       sNoProjet = vbNullString
145       sNoSoumission = rstProjSoum.Fields("IDSoumission")

150       Call rstPiece.Open("SELECT * FROM GRB_Soumission_Pieces WHERE IDSoumission = '" & sNoSoumission & "' AND Type = 'E' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
155     End If

160     Do While Not rstPiece.EOF
165       If rstPiece.Fields("Visible") = True Then
170         bAjouterSection = True
175         bAjouterSousSection = True
180         bAjouterPiece = True

185         rstImpListePiece.CursorLocation = adUseClient

190         Call rstImpListePiece.Open("SELECT * FROM GRB_Impression_ListePiece WHERE IDSection = '" & rstPiece.Fields("IDSection") & "'", g_connData, adOpenDynamic, adLockOptimistic)

195         If Not rstImpListePiece.EOF Then
200           bAjouterSection = False

205           Do While Not rstImpListePiece.EOF
210             If rstImpListePiece.Fields("NomSousSection") = rstPiece.Fields("SousSection") Then
215               bAjouterSousSection = False

220               If rstPiece.Fields("NumItem") <> "Texte" And rstPiece.Fields("NumItem") <> "Text" Then
225                 If rstImpListePiece.Fields("NumItem") = rstPiece.Fields("NumItem") Then
230                   bAjouterPiece = False

235                   rstImpListePiece.Fields("Qté") = Replace(CDbl(rstImpListePiece.Fields("Qté")) + CDbl(rstPiece.Fields("Qté")), ".", ",")

240                   If Not IsNull(rstImpListePiece.Fields("ID")) Then
245                     If rstImpListePiece.Fields("ID") <> "" Then
250                       rstImpListePiece.Fields("ID") = rstImpListePiece.Fields("ID") & ", " & rstPiece.Fields("ID")
255                     Else
260                       rstImpListePiece.Fields("ID") = rstPiece.Fields("ID")
265                     End If
270                   Else
275                     rstImpListePiece.Fields("ID") = rstPiece.Fields("ID")
280                   End If

285                   Call rstImpListePiece.Update

290                   If rstImpListePiece.Fields("Qté") = 0 Then
295                     Call rstImpListePiece.Delete

300                     rstImpListePiece.Filter = "NomSousSection = '" & Replace(rstPiece.Fields("SousSection"), "'", "''") & "'"

305                     If rstImpListePiece.RecordCount = 1 Then
310                       Call rstImpListePiece.Delete

315                       rstImpListePiece.Filter = "IDSection = '" & rstPiece.Fields("IDSection") & "'"

320                       If rstImpListePiece.RecordCount = 1 Then
325                         Call rstImpListePiece.Delete
330                       End If
335                     End If

340                     rstImpListePiece.Filter = ""

345                   End If

350                   Exit Do
355                 End If
360               Else
365                 Exit Do
370               End If
375             End If

380             Call rstImpListePiece.MoveNext
385           Loop
390         End If

395         If bAjouterSection = True Then
400           If m_eLangage = ANGLAIS Then
405             sSection = "NomSectionEN"
410           Else
415             sSection = "NomSectionFR"
420           End If

425           Call rstTemp.Open("SELECT " & sSection & " FROM GRB_SoumProjSectionElec WHERE IDSection = " & rstPiece.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)

              'Ajoute la section dans la liste de pièces
430           Call rstImpListePiece.AddNew

435           rstImpListePiece.Fields("NoLigne") = iCompteurPiece
440           rstImpListePiece.Fields("IDSoumission") = sNoSoumission

445           If Not IsNull(rstTemp.Fields(sSection)) Then
450             rstImpListePiece.Fields("Section") = rstTemp.Fields(sSection)
455          Else
460             rstImpListePiece.Fields("Section") = " "
465           End If

470           rstImpListePiece.Fields("IDSection") = rstPiece.Fields("IDSection")

475           Call rstImpListePiece.Update

480           iCompteurPiece = iCompteurPiece + 1

485           Call rstTemp.Close
490         End If

495         If bAjouterSousSection = True Then
500           sSousSection = rstPiece.Fields("SousSection")

505           If sSousSection = S_PAS_SOUS_SECTION Then
510             sSousSection = " "
515           End If

520           Call rstImpListePiece.AddNew

525           rstImpListePiece.Fields("NoLigne") = iCompteurPiece
530           rstImpListePiece.Fields("IDSoumission") = sNoSoumission
535           rstImpListePiece.Fields("SousSection") = sSousSection
540           rstImpListePiece.Fields("NomSousSection") = rstPiece.Fields("SousSection")
545           rstImpListePiece.Fields("IDSection") = rstPiece.Fields("IDSection")

550           Call rstImpListePiece.Update

555           iCompteurPiece = iCompteurPiece + 1
560         End If

565         If bAjouterPiece = True Then
              'Ajoute la pièce à la liste de pièces
570           Call rstImpListePiece.AddNew

575           rstImpListePiece.Fields("NoLigne") = iCompteurPiece
580           rstImpListePiece.Fields("IDSoumission") = sNoSoumission
585           rstImpListePiece.Fields("NumItem") = rstPiece.Fields("NumItem")
590           rstImpListePiece.Fields("Qté") = rstPiece.Fields("Qté")

595           If m_eLangage = ANGLAIS Then
600             rstImpListePiece.Fields("Description") = rstPiece.Fields("Desc_EN")
605           Else
610             rstImpListePiece.Fields("Description") = rstPiece.Fields("Desc_FR")
615           End If

620           rstImpListePiece.Fields("Manufact") = rstPiece.Fields("Manufact")

625           If m_eType = TYPE_PROJET Then
630             rstImpListePiece.Fields("ID") = rstPiece.Fields("ID")
635           End If

640           rstImpListePiece.Fields("IDSection") = rstPiece.Fields("IDSection")
645           rstImpListePiece.Fields("NomSousSection") = rstPiece.Fields("SousSection")

650           Call rstImpListePiece.Update

655           iCompteurPiece = iCompteurPiece + 1
660         End If

665         Call rstImpListePiece.Close
670       End If

          'Prochaine enregistrement
675       Call rstPiece.MoveNext
680     Loop

        ''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Rapport liste pièce, met dans l'ordre de ligne '
        ''''''''''''''''''''''''''''''''''''''''''''''''''
685     rstImpListePiece.CursorLocation = adUseClient

690     Call rstImpListePiece.Open("SELECT * FROM GRB_impression_Listepiece WHERE TRIM(IDSoumission) = '" & Trim$(sNoSoumission) & "' ORDER BY noligne", g_connData, adOpenDynamic, adLockOptimistic)

695     Set DR_Liste_piece.DataSource = rstImpListePiece

700     Call TraduireImpressionListePiece

        'Affiche la date
705     DR_Liste_piece.Sections("Section3").Controls("lblDate").Caption = ConvertDate(Date)

710     DR_Liste_piece.Sections("Section4").Controls("lblProjet").Caption = sNoProjet

        'Affiche l 'entête
715     DR_Liste_piece.Sections("Section4").Controls("lblSoumission").Caption = sNoSoumission

720     DR_Liste_piece.Sections("Section4").Controls("lblDescription").Caption = rstProjSoum.Fields("Description")

725     Call rstTemp.Open("SELECT NomClient FROM GRB_Client WHERE IDClient = " & rstProjSoum.Fields("IDClient"), g_connData, adOpenDynamic, adLockOptimistic)

730     DR_Liste_piece.Sections("Section4").Controls("lblClient").Caption = rstTemp.Fields("NomClient")

735     Call rstTemp.Close

740     Call rstTemp.Open("SELECT NomContact FROM GRB_Contact WHERE IDContact = " & rstProjSoum.Fields("IDContact"), g_connData, adOpenDynamic, adLockOptimistic)

745     DR_Liste_piece.Sections("Section4").Controls("lblContact").Caption = rstTemp.Fields("nomcontact")

750     Call rstTemp.Close

        'Affiche le rapport liste des pieces
755     DR_Liste_piece.Orientation = rptOrientPortrait

760     Call DR_Liste_piece.Show(vbModal)

765     Call rstImpListePiece.Close
770     Set rstImpListePiece = Nothing

775     Set rstTemp = Nothing

780     Screen.MousePointer = vbDefault

785     Exit Sub

AfficherErreur:

790     woups "frmProjSoumElec", "ImprimerListePieces", Err, Erl
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
  
75      sTotal = "(SUM(" & sDateFin & " - " & sDateDebut & ") * 24) As Total"

80      rstTotal.CursorLocation = adUseServer

85      Call rstTotal.Open("SELECT Type, " & sTotal & " FROM GRB_Punch WHERE " & sFilterNoProjet & " And HeureFin is not null AND HeureDébut is not null GROUP BY Type", g_connData, adOpenDynamic, adLockOptimistic)

90      DR_SoumissionElec.Sections("Section5").Controls("lblTempsDessinReel").Caption = "0"
95      DR_SoumissionElec.Sections("Section5").Controls("lblTempsFabricationReel").Caption = "0"
100     DR_SoumissionElec.Sections("Section5").Controls("lblTempsAssemblageReel").Caption = "0"
105     DR_SoumissionElec.Sections("Section5").Controls("lblTempsProgInterfaceReel").Caption = "0"
110     DR_SoumissionElec.Sections("Section5").Controls("lblTempsProgAutomateReel").Caption = "0"
115     DR_SoumissionElec.Sections("Section5").Controls("lblTempsProgRobotReel").Caption = "0"
120     DR_SoumissionElec.Sections("Section5").Controls("lblTempsVisionReel").Caption = "0"
125     DR_SoumissionElec.Sections("Section5").Controls("lblTempsTestReel").Caption = "0"
130     DR_SoumissionElec.Sections("Section5").Controls("lblTempsInstallationReel").Caption = "0"
135     DR_SoumissionElec.Sections("Section5").Controls("lblTempsMiseServiceReel").Caption = "0"
140     DR_SoumissionElec.Sections("Section5").Controls("lblTempsFormationReel").Caption = "0"
145     DR_SoumissionElec.Sections("Section5").Controls("lblTempsGestionReel").Caption = "0"
150     DR_SoumissionElec.Sections("Section5").Controls("lblTempsShippingReel").Caption = "0"

155     Do While Not rstTotal.EOF
160       If Not IsNull(rstTotal.Fields("Total")) Then
165         Select Case rstTotal.Fields("Type")
              Case "Dessin":        DR_SoumissionElec.Sections("Section5").Controls("lblTempsDessinReel").Caption = Round(rstTotal.Fields("Total"), 2)
170           Case "Fabrication":   DR_SoumissionElec.Sections("Section5").Controls("lblTempsFabricationReel").Caption = Round(rstTotal.Fields("Total"), 2)
175           Case "Assemblage":    DR_SoumissionElec.Sections("Section5").Controls("lblTempsAssemblageReel").Caption = Round(rstTotal.Fields("Total"), 2)
180           Case "ProgInterface": DR_SoumissionElec.Sections("Section5").Controls("lblTempsProgInterfaceReel").Caption = Round(rstTotal.Fields("Total"), 2)
185           Case "ProgAutomate":  DR_SoumissionElec.Sections("Section5").Controls("lblTempsProgAutomateReel").Caption = Round(rstTotal.Fields("Total"), 2)
190           Case "ProgRobot":     DR_SoumissionElec.Sections("Section5").Controls("lblTempsProgRobotReel").Caption = Round(rstTotal.Fields("Total"), 2)
195           Case "Vision":        DR_SoumissionElec.Sections("Section5").Controls("lblTempsVisionReel").Caption = Round(rstTotal.Fields("Total"), 2)
200           Case "Test":          DR_SoumissionElec.Sections("Section5").Controls("lblTempsTestReel").Caption = Round(rstTotal.Fields("Total"), 2)
205           Case "Installation":  DR_SoumissionElec.Sections("Section5").Controls("lblTempsInstallationReel").Caption = Round(rstTotal.Fields("Total"), 2)
210           Case "MiseService":   DR_SoumissionElec.Sections("Section5").Controls("lblTempsMiseServiceReel").Caption = Round(rstTotal.Fields("Total"), 2)
215           Case "Formation":     DR_SoumissionElec.Sections("Section5").Controls("lblTempsFormationReel").Caption = Round(rstTotal.Fields("Total"), 2)
220           Case "Gestion":       DR_SoumissionElec.Sections("Section5").Controls("lblTempsGestionReel").Caption = Round(rstTotal.Fields("Total"), 2)
225           Case "Shipping":      DR_SoumissionElec.Sections("Section5").Controls("lblTempsShippingReel").Caption = Round(rstTotal.Fields("Total"), 2)
230         End Select
235       End If

240       Call rstTotal.MoveNext
245     Loop

250     Call rstTotal.Close
  
255     Call rstTotal.Open("SELECT " & sTotal & " FROM GRB_Punch WHERE " & sFilterNoProjet & " And HeureFin is not null AND HeureDébut is not null", g_connData, adOpenDynamic, adLockOptimistic)

260     If Not IsNull(rstTotal.Fields("Total")) Then
265       DR_SoumissionElec.Sections("Section5").Controls("lblTotalTempsRHReel").Caption = Round(rstTotal.Fields("Total"), 2)
270     Else
275       DR_SoumissionElec.Sections("Section5").Controls("lblTotalTempsRHReel").Caption = "0"
280     End If

285     Call rstTotal.Close
290     Set rstTotal = Nothing

295     Exit Sub

AfficherErreur:

300     woups "frmProjSoumElec", "CalculerTempsReels", Err, Erl
End Sub

Private Sub ImprimerProjSoumFacturation(ByVal rstProjSoum As ADODB.Recordset, ByVal sNoFacture As String)

5       On Error GoTo AfficherErreur

        'Impression de la feuille de soumission
10      Dim rstPiece              As ADODB.Recordset
15      Dim rstTemp               As ADODB.Recordset
20      Dim rstImpProjSoum        As ADODB.Recordset
25      Dim sOrdreSection         As String
30      Dim iCompteurSoum         As Integer
35      Dim sSousSection          As String
40      Dim sSousSectionRS        As String
45      Dim sSection              As String
50      Dim sNoProjet             As String
55      Dim sNoSoumission         As String
60      Dim sCommission           As String
65      Dim sPrixTotal            As String
70      Dim sProfit               As String
75      Dim sTempsFabrication     As String
80      Dim sTotalPiece           As String
85      Dim sImprevue             As String
90      Dim sTotalTemps           As String
95      Dim sManuel               As String
100     Dim dblTotalTemps         As Double
105     Dim dblTempsDessin        As Double
110     Dim dblTempsFabrication   As Double
115     Dim dblTempsAssemblage    As Double
120     Dim dblTempsProgInterface As Double
125     Dim dblTempsProgAutomate  As Double
130     Dim dblTempsProgRobot     As Double
135     Dim dblTempsVision        As Double
140     Dim dblTempsTest          As Double
145     Dim dblTempsInstallation  As Double
150     Dim dblTempsMiseService   As Double
155     Dim dblTempsFormation     As Double
160     Dim dblTempsGestion       As Double
165     Dim dblTempsShipping      As Double
170     Dim dblTotalHebergement   As Double
175     Dim dblTotalRepas         As Double
180     Dim dblTotalTransport     As Double
185     Dim dblTotalUniteMobile   As Double
190     Dim dblPrixEmballage      As Double
195     Dim dblTotalReste         As Double
200     Dim dblTotalAutre         As Double

205     Set rstPiece = New ADODB.Recordset
210     Set rstTemp = New ADODB.Recordset
215     Set rstImpProjSoum = New ADODB.Recordset

        'Supprime les données de l'impression
220     Call g_connData.Execute("DELETE * FROM GRB_impression_soumission")

225     iCompteurSoum = 1
  
230     Screen.MousePointer = vbHourglass

235     Call rstImpProjSoum.Open("SELECT * FROM GRB_impression_soumission", g_connData, adOpenDynamic, adLockOptimistic)
  
        'Ouverture du recordset
240     sNoProjet = rstProjSoum.Fields("IDProjet")
245     sNoSoumission = rstProjSoum.Fields("IDSoumission")

250     Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoProjet & "' AND Type = 'E' AND Facturation = '" & sNoFacture & "' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
            
255     Do While Not rstPiece.EOF
260       sSousSectionRS = rstPiece.Fields("SousSection")
              
265       If sSousSectionRS = S_PAS_SOUS_SECTION Then
270         sSousSectionRS = " "
275       End If
    
280       If sOrdreSection <> rstPiece.Fields("OrdreSection") Then
            'Remplis la table impression_soumission
            'Ajoute seulement la section
285         sOrdreSection = rstPiece.Fields("OrdreSection")
        
290         If m_eLangage = ANGLAIS Then
295           sSection = "NomSectionEN"
300         Else
305           sSection = "NomSectionFR"
310         End If
        
315         Call rstTemp.Open("SELECT " & sSection & " FROM GRB_SoumProjSectionElec WHERE IDSection = " & rstPiece.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
        
            'Ajoute la section dans la soumission
320         Call rstImpProjSoum.AddNew
                
325         rstImpProjSoum.Fields("NoLigne") = iCompteurSoum

330         rstImpProjSoum.Fields("IDSoumission") = sNoProjet
        
335         If Not IsNull(rstTemp.Fields(sSection)) Then
340           rstImpProjSoum.Fields("NomSection") = rstTemp.Fields(sSection)
345         Else
350           rstImpProjSoum.Fields("NomSection") = " "
355         End If
          
360         Call rstImpProjSoum.Update
       
365         iCompteurSoum = iCompteurSoum + 1
       
370         Call rstTemp.Close
        
375         sSousSection = rstPiece.Fields("SousSection")
       
380         If sSousSection = S_PAS_SOUS_SECTION Then
385           sSousSection = " "
390         End If
        
395         Call rstImpProjSoum.AddNew
      
400         rstImpProjSoum.Fields("NoLigne") = iCompteurSoum

405         If m_eType = TYPE_PROJET Then
410           rstImpProjSoum.Fields("IDSoumission") = sNoProjet
415         Else
420           rstImpProjSoum.Fields("IDSoumission") = sNoSoumission
425         End If

430         rstImpProjSoum.Fields("SousSection") = sSousSection
     
435         Call rstImpProjSoum.Update

440         iCompteurSoum = iCompteurSoum + 1
445       Else
            'Ajoute une soussection dans impression_soum
450         If sSousSection <> sSousSectionRS Then
455           sSousSection = sSousSectionRS
        
460           Call rstImpProjSoum.AddNew
      
465           rstImpProjSoum.Fields("NoLigne") = iCompteurSoum

470           rstImpProjSoum.Fields("IDSoumission") = sNoProjet

475           rstImpProjSoum.Fields("SousSection") = sSousSectionRS
                  
480           Call rstImpProjSoum.Update
    
485           iCompteurSoum = iCompteurSoum + 1
490         End If
495       End If
        
          'Ajoute une piece dans impression_soum
500       Call rstImpProjSoum.AddNew
    
505       rstImpProjSoum.Fields("NoLigne") = iCompteurSoum

510       rstImpProjSoum.Fields("IDSoumission") = sNoProjet

515       rstImpProjSoum.Fields("NumItem") = rstPiece.Fields("NumItem")
520       rstImpProjSoum.Fields("Qté") = rstPiece.Fields("Qté")
                   
525       If m_eLangage = ANGLAIS Then
530         rstImpProjSoum.Fields("Description") = rstPiece.Fields("DESC_EN")
535       Else
540         rstImpProjSoum.Fields("Description") = rstPiece.Fields("DESC_FR")
545       End If

        '************************************************************************************************
        'SECTION MODIFIÉ PAR GAÉTAN GINGRAS LE 07 FÉVRIER 2010
        '************************************************************************************************
550       rstImpProjSoum.Fields("MANUFACT") = rstPiece.Fields("MANUFACT")
555       'rstImpProjSoum.Fields("PRIX_LIST") = rstPiece.Fields("PRIX_LIST")

560       'If Not IsNull(rstPiece.Fields("ESCOMPTE")) And Trim(rstPiece.Fields("ESCOMPTE")) <> vbNullString Then
565       '  If rstPiece.Fields("ESCOMPTE") > 0 Then
570       '    rstImpProjSoum.Fields("ESCOMPTE") = Conversion(Replace(rstPiece.Fields("escompte"), ".", ","), MODE_POURCENT)
575       '  Else
580       '    rstImpProjSoum.Fields("ESCOMPTE") = Conversion(Replace(rstPiece.Fields("escompte"), ".", ",") * 100, MODE_POURCENT)
585       '  End If
590       'Else
595       '  rstImpProjSoum.Fields("ESCOMPTE") = Conversion(0, MODE_POURCENT)
600       'End If
   
605       rstImpProjSoum.Fields("PRIX_NET") = rstPiece.Fields("PRIX_NET")
                    
610       Call rstTemp.Open("SELECT NomFournisseur FROM GRB_Fournisseur WHERE IDFRS = " & rstPiece.Fields("IDFRS"), g_connData, adOpenDynamic, adLockOptimistic)
         
615       If Not rstTemp.EOF Then
620         rstImpProjSoum.Fields("NomFournisseur") = rstTemp.Fields("NomFournisseur")
625       End If
        
630       Call rstTemp.Close
              
635       'rstImpProjSoum.Fields("TEMPS") = rstPiece.Fields("TEMPS")
640       'rstImpProjSoum.Fields("TEMPS_TOTAL") = rstPiece.Fields("TEMPS_TOTAL")
645       rstImpProjSoum.Fields("PRIX_TOTAL") = rstPiece.Fields("PRIX_TOTAL")
650       rstImpProjSoum.Fields("PROFIT_ARGENT") = rstPiece.Fields("PROFIT_ARGENT")

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
   
655       Call rstImpProjSoum.Update
    
660       iCompteurSoum = iCompteurSoum + 1
           
          'Prochaine enreg
665       Call rstPiece.MoveNext
670     Loop
      
        'Ferme les tables
675     Call rstImpProjSoum.Close
  
        '''''''''''''''''''''''''''''''''''''''''''''''''
        ' Rapport soumission, met dans l'ordre de ligne '
        '''''''''''''''''''''''''''''''''''''''''''''''''
          
680     Call rstImpProjSoum.Open("SELECT * FROM GRB_impression_soumission WHERE IDSoumission = '" & sNoProjet & "' ORDER BY NoLigne", g_connData, adOpenDynamic, adLockOptimistic)
      
685     Set DR_SoumissionElec.DataSource = rstImpProjSoum

690     Call CalculerPrixFacturation(sNoFacture, sCommission, sPrixTotal, sProfit, sTempsFabrication, sTotalPiece, sImprevue, sTotalTemps, sManuel)
    
695     Call TraduireImpressionSoumission
              
700     DR_SoumissionElec.Sections("Section5").Controls("shpCadrePrixReception").Visible = False
705     DR_SoumissionElec.Sections("Section5").Controls("lblTitrePrixReception").Visible = False
710     DR_SoumissionElec.Sections("Section5").Controls("lblPrixReception").Visible = False

715     DR_SoumissionElec.Sections("Section5").Controls("shpCadrePrixSoumission").Visible = False
720     DR_SoumissionElec.Sections("Section5").Controls("lblTitrePrixSoumission").Visible = False
725     DR_SoumissionElec.Sections("Section5").Controls("lblPrixSoumission").Visible = False

730     DR_SoumissionElec.Sections("Section5").Controls("shpCadreForfait").Visible = False
735     DR_SoumissionElec.Sections("Section5").Controls("lblTitreForfait").Visible = False
740     DR_SoumissionElec.Sections("Section5").Controls("lblForfait").Visible = False

745     DR_SoumissionElec.Sections("Section2").Controls("lblTitreNoFacture").Visible = True
750     DR_SoumissionElec.Sections("Section2").Controls("lblNoFacture").Visible = True

755     DR_SoumissionElec.Sections("Section2").Controls("lblNoFacture").Caption = sNoFacture
     
        'Affiche la date
760     DR_SoumissionElec.Sections("Section3").Controls("lbldate").Caption = ConvertDate(Date)
      
        'Affiche entete
765     DR_SoumissionElec.Sections("Section2").Controls("lblSoumission").Caption = sNoSoumission
          
770     DR_SoumissionElec.Sections("Section2").Controls("lblprojet").Caption = sNoProjet
    
775     DR_SoumissionElec.Sections("Section2").Controls("lbldescription").Caption = rstProjSoum.Fields("Description")
    
780     Call rstTemp.Open("SELECT NomClient FROM GRB_Client WHERE IDClient = " & rstProjSoum.Fields("IDClient"), g_connData, adOpenDynamic, adLockOptimistic)
    
785     DR_SoumissionElec.Sections("Section2").Controls("lblclient").Caption = rstTemp.Fields("NomClient")
    
790     Call rstTemp.Close
    
795     Call rstTemp.Open("SELECT NomContact FROM GRB_Contact WHERE IDContact = " & rstProjSoum.Fields("IDContact"), g_connData, adOpenDynamic, adLockOptimistic)
    
800     DR_SoumissionElec.Sections("Section2").Controls("lblcontact").Caption = rstTemp.Fields("NomContact")
                 
805     Call rstTemp.Close
    
        'affiche pied d'etat
810     If Not IsNull(rstProjSoum.Fields("TauxDessin")) Then
815       DR_SoumissionElec.Sections("Section5").Controls("lblTauxDessin").Caption = rstProjSoum.Fields("TauxDessin")
820     Else
825       DR_SoumissionElec.Sections("Section5").Controls("lblTauxDessin").Caption = "0"
830     End If

835     If Not IsNull(rstProjSoum.Fields("TauxFabrication")) Then
840       DR_SoumissionElec.Sections("Section5").Controls("lblTauxFabrication").Caption = rstProjSoum.Fields("TauxFabrication")
845     Else
850       DR_SoumissionElec.Sections("Section5").Controls("lblTauxFabrication").Caption = "0"
855     End If

860     If Not IsNull(rstProjSoum.Fields("TauxAssemblage")) Then
865       DR_SoumissionElec.Sections("Section5").Controls("lblTauxAssemblage").Caption = rstProjSoum.Fields("TauxAssemblage")
870     Else
875       DR_SoumissionElec.Sections("Section5").Controls("lblTauxAssemblage").Caption = "0"
880     End If

885     If Not IsNull(rstProjSoum.Fields("TauxProgInterface")) Then
890       DR_SoumissionElec.Sections("Section5").Controls("lblTauxProgInterface").Caption = rstProjSoum.Fields("TauxProgInterface")
895     Else
900       DR_SoumissionElec.Sections("Section5").Controls("lblTauxProgInterface").Caption = "0"
905     End If

910     If Not IsNull(rstProjSoum.Fields("TauxProgAutomate")) Then
915       DR_SoumissionElec.Sections("Section5").Controls("lblTauxProgAutomate").Caption = rstProjSoum.Fields("TauxProgAutomate")
920     Else
925       DR_SoumissionElec.Sections("Section5").Controls("lblTauxProgAutomate").Caption = "0"
930     End If

935     If Not IsNull(rstProjSoum.Fields("TauxProgRobot")) Then
940       DR_SoumissionElec.Sections("Section5").Controls("lblTauxProgRobot").Caption = rstProjSoum.Fields("TauxProgRobot")
945     Else
950       DR_SoumissionElec.Sections("Section5").Controls("lblTauxProgRobot").Caption = "0"
955     End If

960     If Not IsNull(rstProjSoum.Fields("TauxVision")) Then
965       DR_SoumissionElec.Sections("Section5").Controls("lblTauxVision").Caption = rstProjSoum.Fields("TauxVision")
970     Else
975       DR_SoumissionElec.Sections("Section5").Controls("lblTauxVision").Caption = "0"
980     End If

985     If Not IsNull(rstProjSoum.Fields("TauxTest")) Then
990       DR_SoumissionElec.Sections("Section5").Controls("lblTauxTest").Caption = rstProjSoum.Fields("TauxTest")
995     Else
1000      DR_SoumissionElec.Sections("Section5").Controls("lblTauxTest").Caption = "0"
1005    End If

1010    If Not IsNull(rstProjSoum.Fields("TauxInstallation")) Then
1015      DR_SoumissionElec.Sections("Section5").Controls("lblTauxInstallation").Caption = rstProjSoum.Fields("TauxInstallation")
1020    Else
1025      DR_SoumissionElec.Sections("Section5").Controls("lblTauxInstallation").Caption = "0"
1030    End If

1035    If Not IsNull(rstProjSoum.Fields("TauxMiseService")) Then
1040      DR_SoumissionElec.Sections("Section5").Controls("lblTauxMiseService").Caption = rstProjSoum.Fields("TauxMiseService")
1045    Else
1050      DR_SoumissionElec.Sections("Section5").Controls("lblTauxMiseService").Caption = "0"
1055    End If

1060    If Not IsNull(rstProjSoum.Fields("TauxFormation")) Then
1065      DR_SoumissionElec.Sections("Section5").Controls("lblTauxFormation").Caption = rstProjSoum.Fields("TauxFormation")
1070    Else
1075      DR_SoumissionElec.Sections("Section5").Controls("lblTauxFormation").Caption = "0"
1080    End If

1085    If Not IsNull(rstProjSoum.Fields("TauxGestion")) Then
1090      DR_SoumissionElec.Sections("Section5").Controls("lblTauxGestion").Caption = rstProjSoum.Fields("TauxGestion")
1095    Else
1100      DR_SoumissionElec.Sections("Section5").Controls("lblTauxGestion").Caption = "0"
1105    End If

1110    If Not IsNull(rstProjSoum.Fields("TauxShipping")) Then
1115      DR_SoumissionElec.Sections("Section5").Controls("lblTauxShipping").Caption = rstProjSoum.Fields("TauxShipping")
1120    Else
1125      DR_SoumissionElec.Sections("Section5").Controls("lblTauxShipping").Caption = "0"
1130    End If

1135    If Not IsNull(rstProjSoum.Fields("TempsDessin")) Then
1140      DR_SoumissionElec.Sections("Section5").Controls("lblTempsDessinSoum").Caption = rstProjSoum.Fields("TempsDessin")
1145    Else
1150      DR_SoumissionElec.Sections("Section5").Controls("lblTempsDessinSoum").Caption = "0"
1155    End If

1160    DR_SoumissionElec.Sections("Section5").Controls("lblTempsFabricationSoum").Caption = sTempsFabrication

1165    If Not IsNull(rstProjSoum.Fields("TempsAssemblage")) Then
1170      DR_SoumissionElec.Sections("Section5").Controls("lblTempsAssemblageSoum").Caption = rstProjSoum.Fields("TempsAssemblage")
1175    Else
1180      DR_SoumissionElec.Sections("Section5").Controls("lblTempsAssemblageSoum").Caption = "0"
1185    End If

1190    If Not IsNull(rstProjSoum.Fields("TempsProgInterface")) Then
1195      DR_SoumissionElec.Sections("Section5").Controls("lblTempsProgInterfaceSoum").Caption = rstProjSoum.Fields("TempsProgInterface")
1200    Else
1205      DR_SoumissionElec.Sections("Section5").Controls("lblTempsProgInterfaceSoum").Caption = "0"
1210    End If

1215    If Not IsNull(rstProjSoum.Fields("TempsProgAutomate")) Then
1220      DR_SoumissionElec.Sections("Section5").Controls("lblTempsProgAutomateSoum").Caption = rstProjSoum.Fields("TempsProgAutomate")
1225    Else
1230      DR_SoumissionElec.Sections("Section5").Controls("lblTempsProgAutomateSoum").Caption = "0"
1235    End If

1240    If Not IsNull(rstProjSoum.Fields("TempsProgRobot")) Then
1245      DR_SoumissionElec.Sections("Section5").Controls("lblTempsProgRobotSoum").Caption = rstProjSoum.Fields("TempsProgRobot")
1250    Else
1255      DR_SoumissionElec.Sections("Section5").Controls("lblTempsProgRobotSoum").Caption = "0"
1260    End If

1265    If Not IsNull(rstProjSoum.Fields("TempsVision")) Then
1270      DR_SoumissionElec.Sections("Section5").Controls("lblTempsVisionSoum").Caption = rstProjSoum.Fields("TempsVision")
1275    Else
1280      DR_SoumissionElec.Sections("Section5").Controls("lblTempsVisionSoum").Caption = "0"
1285    End If

1290    If Not IsNull(rstProjSoum.Fields("TempsTest")) Then
1295      DR_SoumissionElec.Sections("Section5").Controls("lblTempsTestSoum").Caption = rstProjSoum.Fields("TempsTest")
1300    Else
1305      DR_SoumissionElec.Sections("Section5").Controls("lblTempsTestSoum").Caption = "0"
1310    End If

1315    If Not IsNull(rstProjSoum.Fields("TempsInstallation")) Then
1320      DR_SoumissionElec.Sections("Section5").Controls("lblTempsInstallationSoum").Caption = rstProjSoum.Fields("TempsInstallation")
1325    Else
1330      DR_SoumissionElec.Sections("Section5").Controls("lblTempsInstallationSoum").Caption = "0"
1335    End If

1340    If Not IsNull(rstProjSoum.Fields("TempsMiseService")) Then
1345      DR_SoumissionElec.Sections("Section5").Controls("lblTempsMiseServiceSoum").Caption = rstProjSoum.Fields("TempsMiseService")
1350    Else
1355      DR_SoumissionElec.Sections("Section5").Controls("lblTempsMiseServiceSoum").Caption = "0"
1360    End If

1365    If Not IsNull(rstProjSoum.Fields("TempsFormation")) Then
1370      DR_SoumissionElec.Sections("Section5").Controls("lblTempsFormationSoum").Caption = rstProjSoum.Fields("TempsFormation")
1375    Else
1380      DR_SoumissionElec.Sections("Section5").Controls("lblTempsFormationSoum").Caption = "0"
1385    End If

1390    If Not IsNull(rstProjSoum.Fields("TempsGestion")) Then
1395      DR_SoumissionElec.Sections("Section5").Controls("lblTempsGestionSoum").Caption = rstProjSoum.Fields("TempsGestion")
1400    Else
1405      DR_SoumissionElec.Sections("Section5").Controls("lblTempsGestionSoum").Caption = "0"
1410    End If

1415    If Not IsNull(rstProjSoum.Fields("TempsShipping")) Then
1420      DR_SoumissionElec.Sections("Section5").Controls("lblTempsShippingSoum").Caption = rstProjSoum.Fields("TempsShipping")
1425    Else
1430      DR_SoumissionElec.Sections("Section5").Controls("lblTempsShippingSoum").Caption = "0"
1435    End If

1440    If IsNumeric(rstProjSoum.Fields("TempsDessin")) Then
1445      dblTempsDessin = CDbl(rstProjSoum.Fields("TempsDessin"))
1450    Else
1455      dblTempsDessin = 0
1460    End If

1465    If IsNumeric(sTempsFabrication) Then
1470      dblTempsFabrication = CDbl(sTempsFabrication)
1475    Else
1480      dblTempsFabrication = 0
1485    End If

1490    If IsNumeric(rstProjSoum.Fields("TempsAssemblage")) Then
1495      dblTempsAssemblage = CDbl(rstProjSoum.Fields("TempsAssemblage"))
1500    Else
1505      dblTempsAssemblage = 0
1510    End If

1515    If IsNumeric(rstProjSoum.Fields("TempsProgInterface")) Then
1520      dblTempsProgInterface = CDbl(rstProjSoum.Fields("TempsProgInterface"))
1525    Else
1530      dblTempsProgInterface = 0
1535    End If

1540    If IsNumeric(rstProjSoum.Fields("TempsProgAutomate")) Then
1545      dblTempsProgAutomate = CDbl(rstProjSoum.Fields("TempsProgAutomate"))
1550    Else
1555      dblTempsProgAutomate = 0
1560    End If

1565    If IsNumeric(rstProjSoum.Fields("TempsProgRobot")) Then
1570      dblTempsProgRobot = CDbl(rstProjSoum.Fields("TempsProgRobot"))
1575    Else
1580      dblTempsProgRobot = 0
1585    End If

1590    If IsNumeric(rstProjSoum.Fields("TempsVision")) Then
1595      dblTempsVision = CDbl(rstProjSoum.Fields("TempsVision"))
1600    Else
1605      dblTempsVision = 0
1610    End If

1615    If IsNumeric(rstProjSoum.Fields("TempsTest")) Then
1620      dblTempsTest = CDbl(rstProjSoum.Fields("TempsTest"))
1625    Else
1630      dblTempsTest = 0
1635    End If

1640    If IsNumeric(rstProjSoum.Fields("TempsInstallation")) Then
1645      dblTempsInstallation = CDbl(rstProjSoum.Fields("TempsInstallation"))
1650    Else
1655      dblTempsInstallation = 0
1660    End If

1665    If IsNumeric(rstProjSoum.Fields("TempsMiseService")) Then
1670      dblTempsMiseService = CDbl(rstProjSoum.Fields("TempsMiseService"))
1675    Else
1680      dblTempsMiseService = 0
1685    End If

1690    If IsNumeric(rstProjSoum.Fields("TempsFormation")) Then
1695      dblTempsFormation = CDbl(rstProjSoum.Fields("TempsFormation"))
1700    Else
1705      dblTempsFormation = 0
1710    End If

1715    If IsNumeric(rstProjSoum.Fields("TempsGestion")) Then
1720      dblTempsGestion = CDbl(rstProjSoum.Fields("TempsGestion"))
1725    Else
1730      dblTempsGestion = 0
1735    End If

1740    If IsNumeric(rstProjSoum.Fields("TempsShipping")) Then
1745      dblTempsShipping = CDbl(rstProjSoum.Fields("TempsShipping"))
1750    Else
1755      dblTempsShipping = 0
1760    End If

1765    dblTotalTemps = dblTempsDessin + _
                        dblTempsFabrication + _
                        dblTempsAssemblage + _
                        dblTempsProgInterface + _
                        dblTempsProgAutomate + _
                        dblTempsProgRobot + _
                        dblTempsVision + _
                        dblTempsTest + _
                        dblTempsInstallation + _
                        dblTempsMiseService + _
                        dblTempsFormation + _
                        dblTempsGestion + _
                        dblTempsShipping
                          
1770    DR_SoumissionElec.Sections("Section5").Controls("lblTotalTempsRHSoum").Caption = dblTotalTemps

1775    Call CalculerTempsReelsImpression(rstProjSoum.Fields("IDProjet"))

        'Autres frais
1780    If Not IsNull(rstProjSoum.Fields("NbrePersonne")) Then
1785      DR_SoumissionElec.Sections("Section5").Controls("lblNbrePersonne").Caption = rstProjSoum.Fields("NbrePersonne")
1790    Else
1795      DR_SoumissionElec.Sections("Section5").Controls("lblNbrePersonne").Caption = "0"
1800    End If

1805    If Not IsNull(rstProjSoum.Fields("TempsHebergement")) Then
1810      DR_SoumissionElec.Sections("Section5").Controls("lblTempsHebergement").Caption = rstProjSoum.Fields("TempsHebergement")
1815    Else
1820      DR_SoumissionElec.Sections("Section5").Controls("lblTempsHebergement").Caption = "0"
1825    End If

1830    If Not IsNull(rstProjSoum.Fields("TauxHebergement1")) Then
1835      DR_SoumissionElec.Sections("Section5").Controls("lblTauxHebergement1").Caption = rstProjSoum.Fields("TauxHebergement1")
1840    Else
1845      DR_SoumissionElec.Sections("Section5").Controls("lblTauxHebergement1").Caption = "0"
1850    End If

1855    If Not IsNull(rstProjSoum.Fields("TauxHebergement2")) Then
1860      DR_SoumissionElec.Sections("Section5").Controls("lblTauxHebergement2").Caption = rstProjSoum.Fields("TauxHebergement2")
1865    Else
1870      DR_SoumissionElec.Sections("Section5").Controls("lblTauxHebergement2").Caption = "0"
1875    End If

1880    If Not IsNull(rstProjSoum.Fields("TempsRepas")) Then
1885      DR_SoumissionElec.Sections("Section5").Controls("lblTempsRepas").Caption = rstProjSoum.Fields("TempsRepas")
1890    Else
1895      DR_SoumissionElec.Sections("Section5").Controls("lblTempsRepas").Caption = "0"
1900    End If

1905    If Not IsNull(rstProjSoum.Fields("TauxRepas")) Then
1910      DR_SoumissionElec.Sections("Section5").Controls("lblTauxRepas").Caption = rstProjSoum.Fields("TauxRepas")
1915    Else
1920      DR_SoumissionElec.Sections("Section5").Controls("lblTauxRepas").Caption = "0"
1925    End If

1930    If Not IsNull(rstProjSoum.Fields("TempsTransport")) Then
1935      DR_SoumissionElec.Sections("Section5").Controls("lblTempsTransport").Caption = rstProjSoum.Fields("TempsTransport")
1940    Else
1945      DR_SoumissionElec.Sections("Section5").Controls("lblTempsTransport").Caption = "0"
1950    End If

1955    If Not IsNull(rstProjSoum.Fields("TauxTransport")) Then
1960      DR_SoumissionElec.Sections("Section5").Controls("lblTauxTransport").Caption = rstProjSoum.Fields("TauxTransport")
1965    Else
1970      DR_SoumissionElec.Sections("Section5").Controls("lblTauxTransport").Caption = "0"
1975    End If

1980    If Not IsNull(rstProjSoum.Fields("TempsUniteMobile")) Then
1985      DR_SoumissionElec.Sections("Section5").Controls("lblTempsUniteMobile").Caption = rstProjSoum.Fields("TempsUniteMobile")
1990    Else
1995      DR_SoumissionElec.Sections("Section5").Controls("lblTempsUniteMobile").Caption = "0"
2000    End If

2005    If Not IsNull(rstProjSoum.Fields("TauxUniteMobile")) Then
2010      DR_SoumissionElec.Sections("Section5").Controls("lblTauxUniteMobile").Caption = rstProjSoum.Fields("TauxUniteMobile")
2015    Else
2020      DR_SoumissionElec.Sections("Section5").Controls("lblTauxUniteMobile").Caption = "0"
2025    End If

2030    If Not IsNull(rstProjSoum.Fields("PrixEmballage")) Then
2035      DR_SoumissionElec.Sections("Section5").Controls("lblPrixEmballage").Caption = rstProjSoum.Fields("PrixEmballage")
2040    Else
2045      DR_SoumissionElec.Sections("Section5").Controls("lblPrixEmballage").Caption = "0"
2050    End If

2055    DR_SoumissionElec.Sections("Section5").Controls("lblPrixManuel").Caption = rstProjSoum.Fields("Total_Manuel")

2060    DR_SoumissionElec.Sections("Section5").Controls("lblTotalPieceAR").Caption = Conversion(sTotalPiece, MODE_ARGENT)
2065    DR_SoumissionElec.Sections("Section5").Controls("lblImprevue").Caption = Conversion(rstProjSoum.Fields("imprevue"), MODE_POURCENT)
2070    DR_SoumissionElec.Sections("Section5").Controls("lblImprevueAR").Caption = Conversion(sImprevue, MODE_ARGENT)
2075    DR_SoumissionElec.Sections("Section5").Controls("lblTotalTempsAR").Caption = Conversion(sTotalTemps, MODE_ARGENT)
2080    DR_SoumissionElec.Sections("Section5").Controls("lblCommission").Caption = Conversion(rstProjSoum.Fields("commission"), MODE_POURCENT)
2085    DR_SoumissionElec.Sections("Section5").Controls("lblCommissionAR").Caption = Conversion(sCommission, MODE_ARGENT)
2090    DR_SoumissionElec.Sections("Section5").Controls("lblGrandTotalAR").Caption = Conversion(sPrixTotal, MODE_ARGENT)
2095    DR_SoumissionElec.Sections("Section5").Controls("lblProfit").Caption = Conversion(rstProjSoum.Fields("profit") * 100, MODE_POURCENT)
2100    DR_SoumissionElec.Sections("Section5").Controls("lblTotalProfit").Caption = Conversion(sProfit, MODE_ARGENT)

2105    dblTotalHebergement = CDbl(rstProjSoum.Fields("TotalHebergement"))

2110    dblTotalRepas = CDbl(rstProjSoum.Fields("TotalRepas"))

2115    If Not IsNull(rstProjSoum.Fields("TempsTransport")) And Not IsNull(rstProjSoum.Fields("TauxTransport")) Then
2120      dblTotalTransport = CDbl(rstProjSoum.Fields("TempsTransport")) * CDbl(rstProjSoum.Fields("TauxTransport"))
2125    Else
2130      dblTotalTransport = 0
2135    End If

2140    If Not IsNull(rstProjSoum.Fields("TempsUniteMobile")) And Not IsNull(rstProjSoum.Fields("TauxUniteMobile")) Then
2145      dblTotalUniteMobile = CDbl(rstProjSoum.Fields("TempsUniteMobile")) * CDbl(rstProjSoum.Fields("TauxUniteMobile"))
2150    Else
2155      dblTotalUniteMobile = 0
2160    End If

2165    If Not IsNull(rstProjSoum.Fields("PrixEmballage")) Then
2170      dblPrixEmballage = CDbl(Replace(rstProjSoum.Fields("PrixEmballage"), ".", ","))
2175    Else
2180      dblPrixEmballage = 0
2185    End If
        
2190    dblTotalReste = dblTotalHebergement + dblTotalRepas + dblTotalTransport + dblTotalUniteMobile + dblPrixEmballage

2195    dblTotalAutre = dblTotalReste + CDbl(rstProjSoum.Fields("total_manuel"))
    
2200    DR_SoumissionElec.Sections("Section5").Controls("lblAutre").Caption = Conversion(CStr(dblTotalAutre), MODE_ARGENT)

        '************************************************************************************************
        'SECTION MODIFIÉ PAR GAÉTAN GINGRAS LE 07 FÉVRIER 2010
        '************************************************************************************************
        If bFlag = True Then
            DR_SoumissionElec.Sections("Section2").Controls("lbl_DateCommande").Visible = True
            DR_SoumissionElec.Sections("Section2").Controls("lbl_DateReception").Visible = True
            DR_SoumissionElec.Sections("Section1").Controls("txt_DateCommande").Visible = True
            DR_SoumissionElec.Sections("Section1").Controls("txt_DateReception").Visible = True
        Else
            DR_SoumissionElec.Sections("Section2").Controls("lbl_DateCommande").Visible = False
            DR_SoumissionElec.Sections("Section2").Controls("lbl_DateReception").Visible = False
            DR_SoumissionElec.Sections("Section1").Controls("txt_DateCommande").Visible = False
            DR_SoumissionElec.Sections("Section1").Controls("txt_DateReception").Visible = False
        End If
        '************************************************************************************************
        'FIN DE LA SECTION MODIFIÉ
        '************************************************************************************************

        'Affiche le rapport soumission
2205    DR_SoumissionElec.Orientation = rptOrientLandscape
  
2210    Call DR_SoumissionElec.Show(vbModal)
   
2215    Set rstTemp = Nothing
    
2220    Screen.MousePointer = vbDefault

2225    Exit Sub

AfficherErreur:

2230    woups "frmProjSoumElec", "ImprimerProjSoum", Err, Erl
End Sub

Private Sub ImprimerListePiecesFacturation(ByVal rstProjSoum As ADODB.Recordset, ByVal sNoFacture As String)
 
5       On Error GoTo AfficherErreur

        'Impression de la feuille de la liste des pièces
10      Dim rstPiece         As ADODB.Recordset
15      Dim rstTemp          As ADODB.Recordset
20      Dim rstImpListePiece As ADODB.Recordset
25      Dim sOrdreSection    As String
30      Dim iCompteurPiece   As Integer
35      Dim sSousSection     As String
40      Dim sSousSectionRS   As String
45      Dim sSection         As String
50      Dim sNoProjet        As String
55      Dim sNoSoumission    As String

60      Set rstPiece = New ADODB.Recordset
65      Set rstTemp = New ADODB.Recordset
70      Set rstImpListePiece = New ADODB.Recordset

75      Call g_connData.Execute("DELETE * FROM GRB_impression_listepiece")

80      iCompteurPiece = 1
  
85      Screen.MousePointer = vbHourglass
        
90      Call rstImpListePiece.Open("SELECT * FROM GRB_impression_listepiece", g_connData, adOpenDynamic, adLockOptimistic)

95      sOrdreSection = vbNullString

        'Ouverture du recordset
100     sNoProjet = rstProjSoum.Fields("IDProjet")
105     sNoSoumission = rstProjSoum.Fields("IDSoumission")

110     Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoProjet & "' AND Type = 'E' AND Facturation = '" & sNoFacture & "' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
            
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
160             sSection = "NomSectionEN"
165           Else
170             sSection = "NomSectionFR"
175           End If
        
180           Call rstTemp.Open("SELECT " & sSection & " FROM GRB_SoumProjSectionElec WHERE IDSection = " & rstPiece.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
        
              'Ajoute la section dans la liste de pièces
185           Call rstImpListePiece.AddNew
        
190           rstImpListePiece.Fields("NoLigne") = iCompteurPiece
195           rstImpListePiece.Fields("IDSoumission") = sNoSoumission
          
200           If Not IsNull(rstTemp.Fields(sSection)) Then
205             rstImpListePiece.Fields("NomSection") = rstTemp.Fields(sSection)
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
    
                'ajoute une sous-section dans impression_piece
305             Call rstImpListePiece.AddNew
                    
310             rstImpListePiece.Fields("NoLigne") = iCompteurPiece
315             rstImpListePiece.Fields("IDSoumission") = sNoSoumission
320             rstImpListePiece.Fields("SousSection") = sSousSection
            
325             Call rstImpListePiece.Update
            
330             iCompteurPiece = iCompteurPiece + 1
335           End If
340         End If
                       
            'Ajoute la pièce à la liste de pièces
345         Call rstImpListePiece.AddNew
      
350         rstImpListePiece.Fields("NoLigne") = iCompteurPiece
355         rstImpListePiece.Fields("IDSoumission") = sNoSoumission
360         rstImpListePiece.Fields("NumItem") = rstPiece.Fields("NumItem")
365         rstImpListePiece.Fields("Qté") = rstPiece.Fields("Qté")
                
370         If m_eLangage = ANGLAIS Then
375           rstImpListePiece.Fields("Description") = rstPiece.Fields("Desc_EN")
380         Else
385           rstImpListePiece.Fields("Description") = rstPiece.Fields("Desc_FR")
390         End If
          
395         rstImpListePiece.Fields("Manufact") = rstPiece.Fields("Manufact")
        
400         Call rstImpListePiece.Update
    
405         iCompteurPiece = iCompteurPiece + 1
410       End If
          
          'Prochaine enreg
415       Call rstPiece.MoveNext
420     Loop
  
425     Call rstImpListePiece.Close
            
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        ' rapport liste piece, met dans l'ordre de ligne '
        ''''''''''''''''''''''''''''''''''''''''''''''''''
    
430     Call rstImpListePiece.Open("SELECT * FROM GRB_impression_Listepiece WHERE IDSoumission = '" & sNoSoumission & "' ORDER BY noligne", g_connData, adOpenDynamic, adLockOptimistic)
    
435     Set DR_Liste_piece.DataSource = rstImpListePiece
                  
440     Call TraduireImpressionListePiece

445     DR_Liste_piece.Sections("Section4").Controls("lblTitreNoFacture").Visible = True
450     DR_Liste_piece.Sections("Section4").Controls("lblNoFacture").Visible = True

455     DR_Liste_piece.Sections("Section4").Controls("lblNoFacture").Caption = sNoFacture
    
        'Affiche la date
460     DR_Liste_piece.Sections("Section3").Controls("lbldate").Caption = ConvertDate(Date)
    
465     DR_Liste_piece.Sections("Section4").Controls("lblprojet").Caption = sNoProjet
    
        'affiche l 'entête
470     DR_Liste_piece.Sections("Section4").Controls("lblsoumission").Caption = sNoSoumission
 
475     DR_Liste_piece.Sections("section4").Controls("lbldescription").Caption = rstProjSoum.Fields("Description")
    
480     Call rstTemp.Open("SELECT NomClient FROM GRB_Client WHERE IDClient = " & rstProjSoum.Fields("IDClient"), g_connData, adOpenDynamic, adLockOptimistic)
    
485     DR_Liste_piece.Sections("section4").Controls("lblclient").Caption = rstTemp.Fields("NomClient")
  
490     Call rstTemp.Close
   
495     Call rstTemp.Open("SELECT NomContact FROM GRB_Contact WHERE IDContact = " & rstProjSoum.Fields("IDContact"), g_connData, adOpenDynamic, adLockOptimistic)
  
500     DR_Liste_piece.Sections("section4").Controls("lblcontact").Caption = rstTemp.Fields("nomcontact")
    
505     Call rstTemp.Close
 
        'Affiche le rapport liste des pieces
510     DR_Liste_piece.Orientation = rptOrientPortrait

515     Call DR_Liste_piece.Show(vbModal)
      
520     Call rstImpListePiece.Close
525     Set rstImpListePiece = Nothing
      
530     Set rstTemp = Nothing
      
535     Screen.MousePointer = vbDefault

540     Exit Sub

AfficherErreur:

545     woups "frmProjSoumElec", "ImprimerListePieces", Err, Erl
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

100     woups "frmProjSoumElec", "TraduireImpressionListePiece", Err, Erl
End Sub

Private Sub TraduireImpressionSoumission()

5       On Error GoTo AfficherErreur

10      If m_eLangage = ANGLAIS Then
15        If m_eType = TYPE_PROJET Then
20          DR_SoumissionElec.Caption = "Electrical Project"
25          DR_SoumissionElec.Sections("Section2").Controls("lblGrosTitre").Caption = "Electrical Project"
30        Else
35          DR_SoumissionElec.Caption = "Electrical Quote"
40          DR_SoumissionElec.Sections("Section2").Controls("lblGrosTitre").Caption = "Electrical Quote"
45        End If
      
50        DR_SoumissionElec.Sections("Section2").Controls("lblTitreProjet").Caption = "Project :"
55        DR_SoumissionElec.Sections("Section2").Controls("lblTitreSoumission").Caption = "Quote :"
60        DR_SoumissionElec.Sections("Section2").Controls("lblTitreClient").Caption = "Client :"
65        DR_SoumissionElec.Sections("Section2").Controls("lblTitreContact").Caption = "Contact :"

        '************************************************************************************************
        'SECTION MODIFIÉ PAR GAÉTAN GINGRAS LE 07 FÉVRIER 2010
        '************************************************************************************************
     
70        DR_SoumissionElec.Sections("Section2").Controls("lblTitreQuantite").Caption = "Qty"
75        DR_SoumissionElec.Sections("Section2").Controls("lblTitreNoItem").Caption = "Item No."
80        DR_SoumissionElec.Sections("Section2").Controls("lblTitreDescription").Caption = "Description"
85        DR_SoumissionElec.Sections("Section2").Controls("lblTitreManufacturier").Caption = "Manufacturer"
90        'DR_SoumissionElec.Sections("Section2").Controls("lblTitrePrixListe").Caption = "Listed Price"
95        'DR_SoumissionElec.Sections("Section2").Controls("lblTitreEscompte").Caption = "Discount"
100       DR_SoumissionElec.Sections("Section2").Controls("lblTitreCoutant").Caption = "Cost"
105       DR_SoumissionElec.Sections("Section2").Controls("lblTitreFournisseur").Caption = "Supplier"
110       'DR_SoumissionElec.Sections("Section2").Controls("lblTitreTempsMontage").Caption = "Time"
115       'DR_SoumissionElec.Sections("Section2").Controls("lblTitreMontage").Caption = "Fixing Time"
120       DR_SoumissionElec.Sections("Section2").Controls("lblTitreTotal").Caption = "Total"

        'AJOUT PAR GAÉTAN GINGRAS 07 FÉVRIER 2010
        '****************************************************************************************
        DR_SoumissionElec.Sections("Section2").Controls("lbl_DateCommande").Caption = "Order Date"
        DR_SoumissionElec.Sections("Section2").Controls("lbl_DateReception").Caption = "Reception Date"
        '****************************************************************************************
        
125       DR_SoumissionElec.Sections("Section5").Controls("lblTitreDessin").Caption = "Drafting :"
130       DR_SoumissionElec.Sections("Section5").Controls("lblTitreFabrication").Caption = "Manufacturing :"
135       DR_SoumissionElec.Sections("Section5").Controls("lblTitreAssemblage").Caption = "Assembling :"
140       DR_SoumissionElec.Sections("Section5").Controls("lblTitreProgInterface").Caption = "Interface programming :"
145       DR_SoumissionElec.Sections("Section5").Controls("lblTitreProgAutomate").Caption = "PLC programming :"
150       DR_SoumissionElec.Sections("Section5").Controls("lblTitreProgRobot").Caption = "Robot programming :"
155       DR_SoumissionElec.Sections("Section5").Controls("lblTitreVision").Caption = "Vision :"
160       DR_SoumissionElec.Sections("Section5").Controls("lblTitreTest").Caption = "Test :"
165       DR_SoumissionElec.Sections("Section5").Controls("lblTitreInstallation").Caption = "Installation :"
170       DR_SoumissionElec.Sections("Section5").Controls("lblTitreMiseService").Caption = "Activation :"
175       DR_SoumissionElec.Sections("Section5").Controls("lblTitreFormation").Caption = "Training :"
180       DR_SoumissionElec.Sections("Section5").Controls("lblTitreGestion").Caption = "Project management :"
185       DR_SoumissionElec.Sections("Section5").Controls("lblTitreShipping").Caption = "Shipping :"

190       DR_SoumissionElec.Sections("Section5").Controls("lblTitreTauxHoraire").Caption = "Rate / Hours"
195       DR_SoumissionElec.Sections("Section5").Controls("lblTitreTemps").Caption = "Time (Hour)"
200       DR_SoumissionElec.Sections("Section5").Controls("lblTitreTotalPiece").Caption = "Parts Total:"
205       DR_SoumissionElec.Sections("Section5").Controls("lblTitreImprevue").Caption = "Unforeseen:"
210       DR_SoumissionElec.Sections("Section5").Controls("lblTitreTotalTemps").Caption = "Time Total:"
215       DR_SoumissionElec.Sections("Section5").Controls("lblTitreGrandTotal").Caption = "Final Price:"
      
220       DR_SoumissionElec.Sections("Section3").Controls("lblNoPage").Caption = "Page %p of %P"
225       DR_SoumissionElec.Sections("Section5").Controls("lblTitrePrixReception").Caption = "Receiving up to date"
230       DR_SoumissionElec.Sections("Section5").Controls("lblTitrePrixSoumission").Caption = "Quote Price"
235       DR_SoumissionElec.Sections("Section5").Controls("lblTitreForfait").Caption = "Package Deal"
240     Else
245       If m_eType = TYPE_PROJET Then
250         DR_SoumissionElec.Caption = "Projet Électrique"
255         DR_SoumissionElec.Sections("Section2").Controls("lblGrosTitre").Caption = "Projet Électrique"
260       Else
265         DR_SoumissionElec.Caption = "Soumission Électrique"
270         DR_SoumissionElec.Sections("Section2").Controls("lblGrosTitre").Caption = "Soumission Électrique"
275       End If
     
280       DR_SoumissionElec.Sections("Section2").Controls("lblTitreProjet").Caption = "Projet:"
285       DR_SoumissionElec.Sections("Section2").Controls("lblTitreSoumission").Caption = "Soumission:"
290       DR_SoumissionElec.Sections("Section2").Controls("lblTitreClient").Caption = "Client:"
295       DR_SoumissionElec.Sections("Section2").Controls("lblTitreContact").Caption = "Contact:"
      
300       DR_SoumissionElec.Sections("Section2").Controls("lblTitreQuantite").Caption = "Qté"
305       DR_SoumissionElec.Sections("Section2").Controls("lblTitreNoItem").Caption = "No. Item"
310       DR_SoumissionElec.Sections("Section2").Controls("lblTitreDescription").Caption = "Description"
315       DR_SoumissionElec.Sections("Section2").Controls("lblTitreManufacturier").Caption = "Manufacturier"
320       'DR_SoumissionElec.Sections("Section2").Controls("lblTitrePrixListe").Caption = "Prix listé"
325       'DR_SoumissionElec.Sections("Section2").Controls("lblTitreEscompte").Caption = "Escompte"
330       DR_SoumissionElec.Sections("Section2").Controls("lblTitreCoutant").Caption = "Coûtant"
335       DR_SoumissionElec.Sections("Section2").Controls("lblTitreFournisseur").Caption = "Fournisseur"
340       'DR_SoumissionElec.Sections("Section2").Controls("lblTitreTempsMontage").Caption = "Temps"
345       'DR_SoumissionElec.Sections("Section2").Controls("lblTitreMontage").Caption = "Montage"
350       DR_SoumissionElec.Sections("Section2").Controls("lblTitreTotal").Caption = "Total"

        'AJOUT PAR GAÉTAN GINGRAS 07 FÉVRIER 2010
        '****************************************************************************************
        DR_SoumissionElec.Sections("Section2").Controls("lbl_DateCommande").Caption = "Date commandé"
        DR_SoumissionElec.Sections("Section2").Controls("lbl_DateReception").Caption = "Date reçu"
        '****************************************************************************************

        '************************************************************************************************
        'FIN DE LA SECTION DE MODIFICATION
        '************************************************************************************************
        
355       DR_SoumissionElec.Sections("Section5").Controls("lblTitreDessin").Caption = "Dessin :"
360       DR_SoumissionElec.Sections("Section5").Controls("lblTitreFabrication").Caption = "Fabrication :"
365       DR_SoumissionElec.Sections("Section5").Controls("lblTitreAssemblage").Caption = "Assemblage :"
370       DR_SoumissionElec.Sections("Section5").Controls("lblTitreProgInterface").Caption = "Programmation d'interface :"
375       DR_SoumissionElec.Sections("Section5").Controls("lblTitreProgAutomate").Caption = "Programmation d'automate :"
380       DR_SoumissionElec.Sections("Section5").Controls("lblTitreProgRobot").Caption = "Programmation de robot :"
385       DR_SoumissionElec.Sections("Section5").Controls("lblTitreVision").Caption = "Vision :"
390       DR_SoumissionElec.Sections("Section5").Controls("lblTitreTest").Caption = "Test :"
395       DR_SoumissionElec.Sections("Section5").Controls("lblTitreInstallation").Caption = "Installation :"
400       DR_SoumissionElec.Sections("Section5").Controls("lblTitreMiseService").Caption = "Mise en service :"
405       DR_SoumissionElec.Sections("Section5").Controls("lblTitreFormation").Caption = "Formation du personnel :"
410       DR_SoumissionElec.Sections("Section5").Controls("lblTitreGestion").Caption = "Gestion du projet :"
415       DR_SoumissionElec.Sections("Section5").Controls("lblTitreShipping").Caption = "Expédition :"

420       DR_SoumissionElec.Sections("Section5").Controls("lblTitreTauxHoraire").Caption = "Taux Horaire"
425       DR_SoumissionElec.Sections("Section5").Controls("lblTitreTemps").Caption = "Temps (Heure)"
430       DR_SoumissionElec.Sections("Section5").Controls("lblTitreTotalPiece").Caption = "Total pièce:"
435       DR_SoumissionElec.Sections("Section5").Controls("lblTitreImprevue").Caption = "Imprévue:"
440       DR_SoumissionElec.Sections("Section5").Controls("lblTitreTotalTemps").Caption = "Total temps:"
445       DR_SoumissionElec.Sections("Section5").Controls("lblTitreGrandTotal").Caption = "Grand total:"
      
450       DR_SoumissionElec.Sections("Section3").Controls("lblNoPage").Caption = "Page %p de %P"
455       DR_SoumissionElec.Sections("Section5").Controls("lblTitrePrixReception").Caption = "Réception jusqu'à maintenant"
460       DR_SoumissionElec.Sections("Section5").Controls("lblTitrePrixSoumission").Caption = "$ Soumission"
465       DR_SoumissionElec.Sections("Section5").Controls("lblTitreForfait").Caption = "Forfait"
470     End If

475     Exit Sub

AfficherErreur:

480     woups "frmProjSoumElec", "TraduireImpressionSoumission", Err, Erl
End Sub

Private Sub cmdModifier_Click()

5       On Error GoTo AfficherErreur

10      Dim rstProjSoum As ADODB.Recordset
15      Dim sUser       As String

20      Set m_collQteSupp = New Collection
25      Set m_collDateSupp = New Collection
30      Set m_collHeureSupp = New Collection
35      Set m_collNoItemSupp = New Collection

        'Modifier une soumission
40      If cmbProjSoum.ListIndex > -1 Then
45        If Right$(txtNoProjSoum.Text, 2) = "99" Then
50          If m_eType = TYPE_PROJET Then
55            Call MsgBox("Ce projet ne peut pas être modifié!", vbOKOnly, "Erreur")
60          Else
65            Call MsgBox("Cette soumission ne peut pas être modifiée!", vbOKOnly, "Erreur")
70          End If
        
75          Exit Sub
80        End If

85        If m_eType = TYPE_SOUMISSION Then
90          If VerifierSiDejaProjet = True Then
95            Call MsgBox("Vous ne pouvez pas modifier cette soumission, le projet a déjà été créé!", vbOKOnly, "Erreur")

100           Exit Sub
105         End If
110       End If

115       Set rstProjSoum = New ADODB.Recordset

120       Call rstProjSoum.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum ='" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
  
125       If rstProjSoum.Fields("Ouvert") = False Or rstProjSoum.Fields("Verrouillé") = True Then
130         If rstProjSoum.Fields("Ouvert") = False Then
135           If m_eType = TYPE_PROJET Then
140             Call MsgBox("Ce projet est fermé!", vbOKOnly, "Erreur")
145           Else
150             Call MsgBox("Cette soumission est fermée!", vbOKOnly, "Erreur")
155           End If
160         Else
165           If m_eType = TYPE_PROJET Then
170             Call MsgBox("Ce projet est verrouillé!", vbOKOnly, "Erreur")
175           Else
180             Call MsgBox("Cette soumission est verrouillée!", vbOKOnly, "Erreur")
185           End If
190         End If
  
195         Call rstProjSoum.Close
200         Set rstProjSoum = Nothing
  
205         Exit Sub
210       End If
       
215       Call rstProjSoum.Close
220       Set rstProjSoum = Nothing
       
225       If VerifierSiOuvert(sUser) = False Then
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
       
            'On recalcul le prix total
280         Call CalculerPrix
       
285         Call lvwSoumission.Refresh
  
290         Call OuvrirProjSoum(True)
       
295         Screen.MousePointer = vbDefault
300       Else
305         If m_eType = TYPE_PROJET Then
310           Call MsgBox("Ce projet est en modification par " & sUser & "!", vbOKOnly, "Erreur")
315         Else
320           Call MsgBox("Cette soumission est en modification par " & sUser & "!", vbOKOnly, "Erreur")
325         End If
330       End If
335     End If

340     Exit Sub

AfficherErreur:

345     woups "frmProjSoumElec", "cmdModifier_Click", Err, Erl
End Sub

Private Sub cmdsupprimer_Click()

5       On Error GoTo AfficherErreur

10      Dim iReponse    As Integer
15      Dim rstProjSoum As ADODB.Recordset
20      Dim rstProjet   As ADODB.Recordset
25      Dim sSoumission As String
30      Dim sUser       As String
35      Dim iExtension  As Integer

        'Si il y a des enregistrements
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
205           Call MsgBox("Vous ne pouvez pas supprimer cette soumission, le projet a déjà été créé!", vbOKOnly, "Erreur")
        
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

350             Call rstProjet.Open("SELECT IDSoumission FROM GRB_ProjetElec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

355             If Not IsNull(rstProjet.Fields("IDSoumission")) Then
360               sSoumission = rstProjet.Fields("IDSoumission")
365             Else
370               sSoumission = vbNullString
375             End If

380             Call rstProjet.Close
385             Set rstProjet = Nothing

                'Efface les pièces
390             Call g_connData.Execute("DELETE * FROM GRB_Projet_Pieces WHERE IDProjet = '" & txtNoProjSoum.Text & "' AND Type = 'E'")

395             If IsNumeric(Right$(txtNoProjSoum.Text, 2)) Then
400               iExtension = CInt(Right$(txtNoProjSoum.Text, 2))
405             Else
410               iExtension = 0
415             End If

420             If (iExtension >= 60 And iExtension <= 79) Or (iExtension >= 80 And iExtension <= 98) Then
425               Set rstProjSoum = New ADODB.Recordset

430               Call rstProjSoum.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

435               Call g_connData.Execute("DELETE * FROM GRB_Projet_Pieces WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & rstProjSoum.Fields("LiaisonChargeable") & "' AND Provenance = '" & iExtension & "'")

440               Call CalculerTotalRecordset(Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & rstProjSoum.Fields("LiaisonChargeable"))
                  
445               Call rstProjSoum.Close
450               Set rstProjSoum = Nothing
455             End If
        
                'Efface les modifications
460             Call g_connData.Execute("DELETE * FROM GRB_Projet_Modif WHERE IDProjet = '" & txtNoProjSoum.Text & "' AND Type = 'E'")
        
                'Efface la soumission
465             Call g_connData.Execute("DELETE * FROM GRB_ProjetElec WHERE IDProjet = '" & txtNoProjSoum.Text & "'")

                'Efface la soumission dans la table GRB_ProjSoum
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
520             Call g_connData.Execute("DELETE * FROM GRB_Soumission_Pieces WHERE IDSoumission = '" & txtNoProjSoum.Text & "' AND Type = 'E'")
        
                'Efface les modifications
525             Call g_connData.Execute("DELETE * FROM GRB_Soumission_Modif WHERE IDSoumission = '" & txtNoProjSoum.Text & "' AND Type = 'E'")
       
                'Efface la soumission
530             Call g_connData.Execute("DELETE * FROM GRB_SoumissionElec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'")
        
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

625     woups "frmProjSoumElec", "cmdSupprimer_Click", Err, Erl
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

        'Initialisation au mode inactif
35      m_eMode = MODE_INACTIF
  
        'Rempli le combo des clients
40      Call RemplirComboClients(vbNullString)
  
        'Rempli le combo des contacts
45      Call RemplirComboSections
  
        'Rempli le combo des catégories de pièce
50      Call RemplirComboCategoriesPieces
    
55      cmbOuvertFerme.ListIndex = I_CMB_OUVERT
    
60      If m_eType = TYPE_PROJET Then
65        cmbChoix.ListIndex = I_IDX_PROJET
70      Else
75        cmbChoix.ListIndex = I_IDX_SOUMISSION
80      End If

85      Exit Sub

AfficherErreur:

90      woups "frmProjSoumElec", "Form_Load", Err, Erl
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
75        If lvwSoumission.ColumnHeaders.count = 12 Then
80          Exit Sub
85        End If
90      Else
95        m_bDroitPrix = True

          'Si les colonnes sont déjà toute là
100       If lvwSoumission.ColumnHeaders.count = 18 Then
105         Exit Sub
110       End If
115     End If
    
        'Il faut enlever les colonnes avant d'en ajouter d'autres
120     Call lvwSoumission.ColumnHeaders.Clear
      
125     Call lvwSoumission.ColumnHeaders.Add(, , "Qté", 650.1418)
130     Call lvwSoumission.ColumnHeaders.Add(, , "No. Item", 1830.0474)
135     Call lvwSoumission.ColumnHeaders.Add(, , "Description", 3809.746)
140     Call lvwSoumission.ColumnHeaders.Add(, , "Manufacturier", 1154.8348)
    
145     If bCacherPrix = False Then
150       Call lvwSoumission.ColumnHeaders.Add(, , "Prix listé", 920.1261, vbRightJustify)
155       Call lvwSoumission.ColumnHeaders.Add(, , "Escompte", 884.9765, vbRightJustify)
160       Call lvwSoumission.ColumnHeaders.Add(, , "Prix net", 920.1261, vbRightJustify)
165     End If
    
170     Call lvwSoumission.ColumnHeaders.Add(, , "Distributeur", 1005.1655)
175     Call lvwSoumission.ColumnHeaders.Add(, , "Temps", 824.882)
180     Call lvwSoumission.ColumnHeaders.Add(, , "Montage", 824.882)
    
185     If bCacherPrix = False Then
190       Call lvwSoumission.ColumnHeaders.Add(, , "TOTAL", 1099.8426, vbRightJustify)
195       Call lvwSoumission.ColumnHeaders.Add(, , "Profit", 920.1261, vbRightJustify)
200     End If

205     Call lvwSoumission.ColumnHeaders.Add(, , "Commentaire", 1000)

210     If m_eType = TYPE_PROJET Then
215       Call lvwSoumission.ColumnHeaders.Add(, , "ID", 1440)

220       If bCacherPrix = False Then
225         Call lvwSoumission.ColumnHeaders.Add(, , "Facturation", 1440)
230       End If

235       Call lvwSoumission.ColumnHeaders.Add(, , "Date Commande", 1440)
240       Call lvwSoumission.ColumnHeaders.Add(, , "Date Requise", 1440)
245       Call lvwSoumission.ColumnHeaders.Add(, , "Commandé par", 1440)
250       Call lvwSoumission.ColumnHeaders.Add(, , "No Séquentiel", 1440)
255     End If

260     Call lvwSoumission.ColumnHeaders.Add(, , "Provenance", 1440)

265     Exit Sub

AfficherErreur:

270     woups "frmProjSoumElec", "RemplirColonnes", Err, Erl
End Sub

Private Sub BarrerChamps(ByVal bBarrer As Boolean)

5       On Error GoTo AfficherErreur

        'Méthode qui barre ou débarre les champs d'après la variable bBarrer
10      txtProjet.Locked = bBarrer
15      txtNbreManuel.Locked = bBarrer
20      txtPrixManuel.Locked = bBarrer
25      picApprob.Enabled = Not bBarrer
30      txtCheminPhotos.Locked = bBarrer

35      Exit Sub

AfficherErreur:

40      woups "frmProjSoumElec", "BarrerChamps", Err, Erl
End Sub

Private Sub ViderChamps()

5       On Error GoTo AfficherErreur

        'Méthode qui initialise les champs
10      txtClient.Text = vbNullString
15      txtcontact.Text = vbNullString
20      txtProjet.Text = vbNullString
25      txtNbreManuel.Text = 0
30      txtPrixManuel.Text = 0
35      txtTransport.Text = vbNullString
40      txtPrixReception.Text = Conversion("0", MODE_ARGENT)
45      txtPrixSoumission.Text = Conversion("0", MODE_ARGENT)
50      chkCSA.Value = vbUnchecked
55      chkCUL.Value = vbUnchecked
60      chkUL.Value = vbUnchecked
65      chkCUR.Value = vbUnchecked
70      chkUR.Value = vbUnchecked
75      chkCE.Value = vbUnchecked
80      txtPrixTotal.Text = 0
85      txtProfit.Text = 0
90      txtDelais.Text = vbNullString
95      txtCommission.Text = 0
100     txtNoSoumission.Text = vbNullString
105     txtCheminPhotos.Text = vbNullString
110     txtForfait.Text = vbNullString
115     lblForfaitInitiale.Caption = vbNullString
    
120     cmbtransport.ListIndex = I_TRANS_FAB_GRANBY

125     cmbclient.ListIndex = -1

130     m_bSansTemps = False
135     lblPasTemps.Visible = False
140     tmrTemps.Enabled = False
  
145     Call lvwSoumission.ListItems.Clear

150     Exit Sub

AfficherErreur:

155     woups "frmProjSoumElec", "ViderChamps", Err, Erl
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
35          Call rstProjSoum.Open("SELECT IDProjet FROM GRB_ProjetElec INNER JOIN GRB_ProjSoum ON GRB_ProjetElec.IDProjet = GRB_ProjSoum.IDProjSoum WHERE Ouvert = True ORDER BY IDProjet DESC", g_connData, adOpenDynamic, adLockOptimistic)
40        Else
45          Call rstProjSoum.Open("SELECT IDProjet FROM GRB_ProjetElec ORDER BY IDProjet DESC", g_connData, adOpenDynamic, adLockOptimistic)
50        End If
55      Else
60        If cmbOuvertFerme.ListIndex = I_CMB_OUVERT Then
65          Call rstProjSoum.Open("SELECT IDSoumission FROM GRB_SoumissionElec INNER JOIN GRB_ProjSoum ON GRB_SoumissionElec.IDSoumission = GRB_ProjSoum.IDProjSoum WHERE Ouvert = True ORDER BY IDSoumission DESC", g_connData, adOpenDynamic, adLockOptimistic)
70        Else
75          Call rstProjSoum.Open("SELECT IDSoumission FROM GRB_SoumissionElec ORDER BY IDSoumission DESC", g_connData, adOpenDynamic, adLockOptimistic)
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

180     woups "frmProjSoumElec", "RemplirComboProjSoum", Err, Erl
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
  
70      rstPieceFRS.CursorLocation = adUseServer
  
75      Call rstPieceFRS.Open("SELECT PrixReel, PRIX_NET, PRIX_SP, DeviseMonétaire FROM GRB_PiecesFRS WHERE PIECE = '" & Replace(sNoItem, "'", "''") & "' AND Type = 'E'", g_connData, adOpenDynamic, adLockOptimistic)
    
80      Do While Not rstPieceFRS.EOF
85        If rstPieceFRS.Fields("PRIX_NET") <> vbNullString Then
90          sPrixCalcul = rstPieceFRS.Fields("PRIX_NET")
95        Else
100         If rstPieceFRS.Fields("PRIX_SP") <> vbNullString Then
105           sPrixCalcul = rstPieceFRS.Fields("PRIX_SP")
110         End If
115       End If

120       sPrixCalcul = Replace(sPrixCalcul, ".", ",")

125       If rstPieceFRS.Fields("DeviseMonétaire") = "USA" Then
130         rstPieceFRS.Fields("PrixReel") = Conversion(CStr(Round(CDbl(sPrixCalcul) / CDbl(sTauxUSA), 4)), MODE_DECIMAL, 4)
135       Else
140         If rstPieceFRS.Fields("DeviseMonétaire") = "SPA" Then
145           rstPieceFRS.Fields("PrixReel") = Conversion(CStr(Round(CDbl(sPrixCalcul) / CDbl(sTauxSPA), 4)), MODE_DECIMAL, 4)
150         Else
155           rstPieceFRS.Fields("PrixReel") = sPrixCalcul
160         End If
165       End If

170       Call rstPieceFRS.Update

175       Call rstPieceFRS.MoveNext
180     Loop

185     Call rstPieceFRS.Close
190     Set rstPieceFRS = Nothing

195     Exit Sub

AfficherErreur:

200     woups "frmProjSoumElec", "CalculerPrixReel", Err, Erl
End Sub

Private Sub RemplirListViewFournisseur()

5       On Error GoTo AfficherErreur

        'Rempli le listview des distributeurs pour une pièce choisie
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
  
        'vide le lister
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
150       Call rstPieceFRS.Open("SELECT GRB_PiecesFRS.*, GRB_Fournisseur.NomFournisseur FROM GRB_PiecesFRS INNER JOIN GRB_Fournisseur ON GRB_PiecesFRS.IDFRS = GRB_Fournisseur.IDFRS WHERE PIECE = '" & Trim$(Replace(lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE), "'", "''")) & "' AND Type = 'E' ORDER BY CDbl(PrixReel)", g_connData, adOpenDynamic, adLockOptimistic)
155     Else
160       If m_bRecherchePiece = True Then
165         Call rstPieceFRS.Open("SELECT GRB_PiecesFRS.*, GRB_Fournisseur.NomFournisseur FROM GRB_PiecesFRS INNER JOIN GRB_Fournisseur ON GRB_PiecesFRS.IDFRS = GRB_Fournisseur.IDFRS WHERE PIECE = '" & Trim$(Replace(lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM), "'", "''")) & "' AND Type = 'E' ORDER BY CDbl(PrixReel)", g_connData, adOpenDynamic, adLockOptimistic)
170       Else
175         Call rstPieceFRS.Open("SELECT GRB_PiecesFRS.*, GRB_Fournisseur.NomFournisseur FROM GRB_PiecesFRS INNER JOIN GRB_Fournisseur ON GRB_PiecesFRS.IDFRS = GRB_Fournisseur.IDFRS WHERE PIECE = '" & Trim$(Replace(lvwPieces.SelectedItem.SubItems(I_COL_PIECES_NO_ITEM), "'", "''")) & "' AND Type = 'E' ORDER BY CDbl(PrixReel)", g_connData, adOpenDynamic, adLockOptimistic)
180       End If
185     End If

        'Tant qu'il y a des fournisseur de la pièce, on ajoute dans le ListView
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
          'CAN = noir, USA ou ESP = bleu
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
345       If Not IsNull(rstPieceFRS.Fields("PERS_RESS")) Then
350         If Trim(rstPieceFRS.Fields("PERS_RESS")) <> vbNullString Then
355           Call rstContact.Open("SELECT NomContact FROM GRB_Contact WHERE IDContact = " & rstPieceFRS.Fields("PERS_RESS"), g_connData, adOpenDynamic, adLockOptimistic)
        
360           If Not rstContact.EOF Then
365             itmFRS.SubItems(I_COL_FRS_PERS_RESS) = rstContact.Fields("NomContact")

370             itmFRS.ListSubItems(I_COL_FRS_PERS_RESS).ForeColor = lColor
375           End If

380           Call rstContact.Close
385         End If
390       End If
                                          
          'Date
395       If Not IsNull(rstPieceFRS.Fields("Date")) Then
400         itmFRS.SubItems(I_COL_FRS_DATE) = rstPieceFRS.Fields("Date")
405       Else
410         itmFRS.SubItems(I_COL_FRS_DATE) = vbNullString
415       End If

420       itmFRS.ListSubItems(I_COL_FRS_DATE).ForeColor = lColor
                          
          'Entrer par
425       If Not IsNull(rstPieceFRS.Fields("Entrer_Par")) Then
430         itmFRS.SubItems(I_COL_FRS_ENTRER_PAR) = rstPieceFRS.Fields("ENTRER_PAR")
435       Else
440         itmFRS.SubItems(I_COL_FRS_ENTRER_PAR) = vbNullString
445       End If

450       itmFRS.ListSubItems(I_COL_FRS_ENTRER_PAR).ForeColor = lColor
                                 
          'Valide
455       If Not IsNull(rstPieceFRS.Fields("Valide")) Then
460         itmFRS.SubItems(I_COL_FRS_VALIDE) = rstPieceFRS.Fields("Valide")
465       Else
470         itmFRS.SubItems(I_COL_FRS_VALIDE) = vbNullString
475       End If
                            
480       itmFRS.ListSubItems(I_COL_FRS_VALIDE).ForeColor = lColor
                             
          'Prix listé
485       If rstPieceFRS.Fields("PRIX_LIST") <> vbNullString Then
490         itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = Conversion(Round(Replace(rstPieceFRS.Fields("PRIX_LIST"), ".", ","), 4), MODE_ARGENT, 4)
495       End If

500       itmFRS.ListSubItems(I_COL_FRS_PRIX_LIST).ForeColor = lColor
                             
          'Escompte
505       If rstPieceFRS.Fields("ESCOMPTE") <> vbNullString Then
510         itmFRS.SubItems(I_COL_FRS_ESCOMPTE) = Conversion(Replace(Replace(rstPieceFRS.Fields("ESCOMPTE"), "_", vbNullString), ".", ",") * 100, MODE_POURCENT)
515       End If

520       itmFRS.ListSubItems(I_COL_FRS_ESCOMPTE).ForeColor = lColor

          'Prix net
525       If rstPieceFRS.Fields("PRIX_NET") <> vbNullString Then
530         itmFRS.SubItems(I_COL_FRS_PRIX_NET) = Conversion(Round(Replace(rstPieceFRS.Fields("PRIX_NET"), ".", ","), 4), MODE_ARGENT, 4)
535       End If

540       itmFRS.ListSubItems(I_COL_FRS_PRIX_NET).ForeColor = lColor
      
          'Prix spécial
545       If rstPieceFRS.Fields("PRIX_SP") <> vbNullString Then
550         itmFRS.SubItems(I_COL_FRS_PRIX_SP) = Conversion(Round(rstPieceFRS.Fields("PRIX_SP"), 4), MODE_ARGENT, 4)
555       End If

560       itmFRS.ListSubItems(I_COL_FRS_PRIX_SP).ForeColor = lColor

          'Quoter
565       If rstPieceFRS.Fields("QUOTER") = True Then
570         itmFRS.SubItems(I_COL_FRS_QUOTER) = "Oui"
575       Else
580         itmFRS.SubItems(I_COL_FRS_QUOTER) = "Non"
585       End If

590       itmFRS.ListSubItems(I_COL_FRS_QUOTER).ForeColor = lColor

595       If rstPieceFRS.Fields("IDFRS") = 717 Then 'Si le fournisseur est "SOLUTION GRB Inc."
600         Set rstInv = New ADODB.Recordset

605         Call rstInv.Open("SELECT * FROM GRB_InventaireElec WHERE TRIM(NoItem) = '" & Trim(rstPieceFRS.Fields("PIECE")) & "'", g_connData, adOpenForwardOnly, adLockReadOnly)
        
610         If Not rstInv.EOF Then
615           If Not IsNull(rstInv.Fields("QuantitéStock")) Then
620             itmFRS.SubItems(I_COL_FRS_STOCK) = rstInv.Fields("QuantitéStock")
625           Else
630             itmFRS.SubItems(I_COL_FRS_STOCK) = 0
635           End If
640         End If

645         Call rstInv.Close
650         Set rstInv = Nothing
655       End If

          'Pour garder en mémoire le prix d'origine, je le mets dans le
          'tag de la colonne Prix Listé
660       If itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = vbNullString Then
665         itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = " "
670       End If
   
675       If rstPieceFRS.Fields("PRIX_NET") <> vbNullString Then
680         If rstPieceFRS.Fields("PRIX_LIST") = "0,00" Or rstPieceFRS.Fields("PRIX_LIST") = "0" Then
685           itmFRS.ListSubItems(I_COL_FRS_PRIX_LIST).Tag = Replace(rstPieceFRS.Fields("PRIX_NET"), ".", ",")
690         Else
695           itmFRS.ListSubItems(I_COL_FRS_PRIX_LIST).Tag = Replace(rstPieceFRS.Fields("PRIX_LIST"), ".", ",")
700         End If
705       Else
710         itmFRS.ListSubItems(I_COL_FRS_PRIX_LIST).Tag = Replace(rstPieceFRS.Fields("PRIX_SP"), ".", ",")
715       End If

720       If itmFRS.SubItems(I_COL_FRS_ENTRER_PAR) = vbNullString Then
725         itmFRS.SubItems(I_COL_FRS_ENTRER_PAR) = " "
730       End If

735       itmFRS.ListSubItems(I_COL_FRS_ENTRER_PAR).Tag = rstPieceFRS.Fields("NoEnreg")

740       If itmFRS.SubItems(I_COL_FRS_PERS_RESS) = "" Then
745         itmFRS.SubItems(I_COL_FRS_PERS_RESS) = " "
750       End If

755       itmFRS.ListSubItems(I_COL_FRS_PERS_RESS).Tag = sDevise
  
760       Call rstPieceFRS.MoveNext
765     Loop
   
        'Ferme la table
770     Call rstPieceFRS.Close
775     Set rstPieceFRS = Nothing

780     Set rstContact = Nothing

785     If m_bPieceInutile = False Then
790       If lvwSoumission.ListItems.count > 0 Then
795         If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor <> COLOR_BRUN Then
800           bAjouterDP = True
805         Else
810           If m_bChangementFRS = False Then
815             bAjouterDP = True
820           End If
825         End If
830       Else
835         bAjouterDP = True
840       End If
845     Else
850       If m_bChangementFRS = True Then
855         If lvwSoumission.ListItems.count > 0 Then
860           If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor <> COLOR_BRUN Then
865             bAjouterDP = True
870           End If
875         Else
880           bAjouterDP = True
885         End If
890       End If
895     End If

900     If bAjouterDP = True Then
905       Set itmFRS = lvwfournisseur.ListItems.Add

910       itmFRS.Text = "CHOISIR ULTÉRIEUREMENT"

915       itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = " "
920       itmFRS.SubItems(I_COL_FRS_ESCOMPTE) = " "
925       itmFRS.SubItems(I_COL_FRS_PRIX_NET) = " "
930       itmFRS.SubItems(I_COL_FRS_PRIX_SP) = " "
935     End If

940     Exit Sub

AfficherErreur:

945     woups "frmProjSoumElec", "RemplirListViewFournisseur", Err, Erl
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
      
95      Call rstPieces.Open("SELECT * FROM GRB_CatalogueElec WHERE CATEGORIE = '" & sCategorie & "' ORDER BY " & sOrderBy, g_connData, adOpenDynamic, adLockOptimistic)

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
135                  bDebut = True
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
        
            'TEMPS
275         If Not IsNull(rstPieces.Fields("TEMPS")) Then
280           itmPieces.Tag = Trim(rstPieces.Fields("TEMPS"))
285         Else
290           itmPieces.Tag = vbNullString
295         End If
                    
            'PIECE_GRB
300         If Not IsNull(rstPieces.Fields("PIECE_GRB")) Then
305           itmPieces.Text = Trim(rstPieces.Fields("PIECE_GRB"))
310         Else
315           itmPieces.Text = vbNullString
320         End If
        
            'PIECE
325         If Not IsNull(rstPieces.Fields("PIECE")) Then
330           itmPieces.SubItems(I_COL_PIECES_NO_ITEM) = Trim(rstPieces.Fields("PIECE"))
335         Else
340           itmPieces.SubItems(I_COL_PIECES_NO_ITEM) = vbNullString
345         End If
        
            'FABRICANT
350         If Not IsNull(rstPieces.Fields("FABRICANT")) Then
355           itmPieces.SubItems(I_COL_PIECES_MANUFACT) = Trim(rstPieces.Fields("FABRICANT"))
360         Else
365           itmPieces.SubItems(I_COL_PIECES_MANUFACT) = vbNullString
370         End If
                   
            'DESCR_FR
375         If Not IsNull(rstPieces.Fields("DESC_FR")) Then
380           itmPieces.SubItems(I_COL_PIECES_DESCR_FR) = Trim(rstPieces.Fields("DESC_FR"))
385         Else
390           itmPieces.SubItems(I_COL_PIECES_DESCR_FR) = vbNullString
395         End If
        
            'DESCR_EN
400         If Not IsNull(rstPieces.Fields("DESC_EN")) Then
405           itmPieces.SubItems(I_COL_PIECES_DESCR_EN) = Trim(rstPieces.Fields("DESC_EN"))
410         Else
415           itmPieces.SubItems(I_COL_PIECES_DESCR_EN) = vbNullString
420         End If
425       End If
        
430       Call rstPieces.MoveNext
435     Loop
      
440     Call rstPieces.Close
445     Set rstPieces = Nothing
  
450     Exit Sub

AfficherErreur:

455     woups "frmProjSoumElec", "RemplirListViewPieces", Err, Erl
End Sub

Private Function TrouverIndexSection(ByVal sSousSection As String) As Integer

5       On Error GoTo AfficherErreur

        'recherche la section et l'ajouter si elle n'a pas été trouvée
10      Dim iCompteur         As Integer
15      Dim iIndex            As Integer
20      Dim iTagSection       As Integer
25      Dim iIDSection        As Integer
30      Dim iIndexSect        As Integer
35      Dim bTrouverSect      As Boolean
40      Dim bTrouverSSect     As Boolean
45      Dim bTrouverIndexItem As Boolean
50      Dim iIndexSSection    As Integer
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
100       For iCompteur = 1 To lvwSoumission.ListItems.count
            'Si c'est écrit le nom de la section dans la colonne Piece
105          If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) = cmbSections.Text Then
              'La section a été trouvée
110           bTrouverSect = True
                            
              'On stock l'index de la section
115           iIndexSect = iCompteur
        
              'On commence à rechercher la sous-section à l'index suivant
120           iCompteur = iCompteur + 1
        
              'Tant que le tag du listItem est égal à l'index de la section
125           Do While lvwSoumission.ListItems(iCompteur).Tag = cmbSections.ItemData(cmbSections.ListIndex)
                'Si c'est écrit le nom de la sous-section dans la colonne Description
130             If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_DESCR) = sSousSection Then
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
    
410       Call ValeurParDefaut(itmSoum)
    
415       iIndex = iIndex + 1
    
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

520     woups "frmProjSoumElec", "TrouverIndexSection", Err, Erl
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

180     woups "frmProjSoumElec", "AjouterSousSection", Err, Erl
End Function

Private Sub AjouterNegatifDansListView(ByVal dblQuantite As Double, ByVal sSousSection As String)

5       On Error GoTo AfficherErreur

10      Dim iIndex      As Integer
15      Dim itmSoum     As ListItem
20      Dim iCompteur   As Integer
25      Dim iIDSection  As Integer
30      Dim iTagSection As Integer
35      Dim bSelected   As Boolean
40      Dim iIndexSel   As Integer
45      Dim dblTempsMec As Double
50      Dim lColor      As Long
55      Dim rstProjet   As ADODB.Recordset
60      Dim bQteOK      As Boolean
65      Dim sNoProjet   As String
70      Dim sPrixList   As String
75      Dim sEscompte   As String
80      Dim sPrixNet    As String
85      Dim sTemps      As String
90      Dim dblTotalQte As Double

95      Set rstProjet = New ADODB.Recordset
  
100     If Right$(txtNoProjSoum.Text, 2) >= 60 And Right$(txtNoProjSoum.Text, 2) <= 98 Then
105       sNoProjet = Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & m_sLiaison

110       If m_bRecherchePiece = True Then
115         Call rstProjet.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoProjet & "' AND NumItem = '" & Replace(lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM), "'", "''") & "' AND IDFRS = " & lvwfournisseur.SelectedItem.Tag, g_connData, adOpenDynamic, adLockOptimistic)
120       Else
125         Call rstProjet.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoProjet & "' AND NumItem = '" & Replace(lvwPieces.SelectedItem.SubItems(I_COL_PIECES_NO_ITEM), "'", "''") & "' AND IDFRS = " & lvwfournisseur.SelectedItem.Tag, g_connData, adOpenDynamic, adLockOptimistic)
130       End If
135     End If

140     If Not rstProjet.EOF Then
145       Do While Not rstProjet.EOF
150         dblTotalQte = dblTotalQte + rstProjet.Fields("Qté")

155         Call rstProjet.MoveNext
160       Loop

165       If dblTotalQte >= Abs(dblQuantite) Then
170         bQteOK = True
175       End If
180     Else
185       Call MsgBox("La pièce n'existe pas dans le projet " & sNoProjet, vbOKOnly, "Erreur")

190       Call rstProjet.Close
195       Set rstProjet = Nothing

200       Exit Sub
205     End If
  
210     If bQteOK = True Then
215       Call rstProjet.MovePrevious

220       sPrixList = rstProjet.Fields("Prix_List")
225       sEscompte = rstProjet.Fields("Escompte")
230       sPrixNet = rstProjet.Fields("Prix_Net")
235       sTemps = rstProjet.Fields("Temps")
240     Else
245       If m_bRecherchePiece = True Then
250         Call MsgBox("Il n'y a pas assez de " & lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM) & " dans le projet " & sNoProjet & " pour en enlever " & Abs(dblQuantite) & "!", vbOKOnly, "Erreur")
255       Else
260         Call MsgBox("Il n'y a pas assez de " & lvwPieces.SelectedItem.SubItems(I_COL_PIECES_NO_ITEM) & " dans le projet " & sNoProjet & " pour en enlever " & Abs(dblQuantite) & "!", vbOKOnly, "Erreur")
265       End If

270       Call rstProjet.Close
275       Set rstProjet = Nothing

280       Exit Sub
285     End If

290     Call rstProjet.Close
295     Set rstProjet = Nothing
  
300     bSelected = False
  
        'S'il y a des items dans le ListView
305     If lvwSoumission.ListItems.count > 0 Then
          'Si ce n'est pas le premier qui est sélectionné
          '(le premier est sélectionné par défaut)
310       If lvwSoumission.SelectedItem.Index > 1 Then
315         bSelected = True

320         iIndexSel = lvwSoumission.SelectedItem.Index
325       End If
330     End If

        'si le premier est sélectionné, on le considère comme si aucun n'est sélectionné
335     If bSelected = False Then
340       iIndex = TrouverIndexSection(sSousSection)
345     Else
          'Sinon, on l'ajoute à l'endroit sélectionné
350       iIndex = iIndexSel
355     End If

360     Set itmSoum = lvwSoumission.ListItems.Add(iIndex)

365     itmSoum.Checked = True

        'Quantité
370     itmSoum.Text = dblQuantite

375     If lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_QUOTER) = "Oui" Then
380       itmSoum.Text = itmSoum.Text & "*"
385       itmSoum.ForeColor = COLOR_VERT
390       itmSoum.Bold = True
395     Else
400       itmSoum.ForeColor = COLOR_NOIR
405       itmSoum.Bold = False
410     End If

        'On met l'id de la section dans le tag du listItem
415     itmSoum.Tag = cmbSections.ItemData(cmbSections.ListIndex) 'IDSection
                                                                                                         
        'No d'item
420     If m_bRecherchePiece = True Then
425       itmSoum.SubItems(I_COL_SOUM_PIECE) = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM)
430     Else
435       itmSoum.SubItems(I_COL_SOUM_PIECE) = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_NO_ITEM)
440     End If

445     itmSoum.ListSubItems(I_COL_SOUM_PIECE).ForeColor = lColor
  
        'On met le nom de la sous-section dans le tag du no d'item
450     itmSoum.ListSubItems(I_COL_SOUM_PIECE).Tag = sSousSection
  
        'On met la description en francais dans la colonne et la description en anglais
        'dans le tag
455     If m_eLangage = ANGLAIS Then
460       If m_bRecherchePiece = True Then
465         itmSoum.SubItems(I_COL_SOUM_DESCR) = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_DESCR_EN)
470         itmSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_DESCR_FR)
475       Else
480         itmSoum.SubItems(I_COL_SOUM_DESCR) = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_DESCR_EN)
485         itmSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_DESCR_FR)
490       End If
495     Else
500       If m_bRecherchePiece = True Then
505         itmSoum.SubItems(I_COL_SOUM_DESCR) = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_DESCR_FR)
510         itmSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_DESCR_EN)
515       Else
520         itmSoum.SubItems(I_COL_SOUM_DESCR) = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_DESCR_FR)
525         itmSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_DESCR_EN)
530       End If
535     End If

540     itmSoum.ListSubItems(I_COL_SOUM_DESCR).ForeColor = lColor
          
        'On met le nom du fabricant dans la colonne et l'ordre de la section dans le tag
545     If m_bRecherchePiece = True Then
550       itmSoum.SubItems(I_COL_SOUM_MANUFACT) = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_MANUFACT)
555     Else
560       itmSoum.SubItems(I_COL_SOUM_MANUFACT) = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_MANUFACT)
565     End If

570     itmSoum.ListSubItems(I_COL_SOUM_MANUFACT).Tag = cmbSections.ListIndex + 1 'Ordre de la section
  
575     itmSoum.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = lColor

        'Prix listé
580     If Trim$(sPrixList) = vbNullString Then
585       itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion("0", MODE_ARGENT, 4)
590     Else
595       itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(sPrixList, MODE_ARGENT, 4)
600       itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = sPrixList
605     End If
       
610     itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = lColor
       
        'Si il y a un prix net, on le met l'escompte et le prix net sinon, on prend le prix
        'spécial pour mettre dans le prix net
615     If Trim$(sEscompte) <> vbNullString Then
620       itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = sEscompte
625     Else
630       itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion("0", MODE_POURCENT)
635     End If
      
640     itmSoum.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = lColor

645     If Trim$(sPrixNet) <> vbNullString Then
650       itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(sPrixNet, MODE_ARGENT, 4)
655     Else
660       itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion("0", MODE_ARGENT, 4)
665     End If

670     itmSoum.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = lColor

675     itmSoum.SubItems(I_COL_SOUM_DISTRIB) = lvwfournisseur.SelectedItem.Text
680     itmSoum.ListSubItems(I_COL_SOUM_DISTRIB).Tag = lvwfournisseur.SelectedItem.Tag
   
685     itmSoum.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = lColor
    
        'Temps
690     itmSoum.SubItems(I_COL_SOUM_TEMPS) = sTemps

695     itmSoum.ListSubItems(I_COL_SOUM_TEMPS).ForeColor = lColor
    
        'Si le temps n'est pas vide
700     If Trim$(itmSoum.SubItems(I_COL_SOUM_TEMPS)) <> vbNullString Then
          'On calcul le temps * quantité pour la colonne montage
705       itmSoum.SubItems(I_COL_SOUM_MONTAGE) = CDbl(Replace(itmSoum.SubItems(I_COL_SOUM_TEMPS), ".", ",")) * CDbl(Replace(itmSoum.Text, "*", vbNullString))
710     Else
715       itmSoum.SubItems(I_COL_SOUM_MONTAGE) = vbNullString
720     End If
      
725     itmSoum.ListSubItems(I_COL_SOUM_MONTAGE).ForeColor = lColor
      
        'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
730     itmSoum.SubItems(I_COL_SOUM_TOTAL) = Conversion(Round(Replace(itmSoum.Text, "*", vbNullString) * itmSoum.SubItems(I_COL_SOUM_PRIX_NET) * CSng(m_sProfit), 2), MODE_ARGENT, 2)
      
735     itmSoum.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = lColor
      
        'Pour le profit, c'est le prix total - (prix net * quantité)
740     itmSoum.SubItems(I_COL_SOUM_PROFIT) = Conversion(CStr(Round(itmSoum.SubItems(I_COL_SOUM_TOTAL) - (itmSoum.SubItems(I_COL_SOUM_PRIX_NET) * Replace(itmSoum.Text, "*", vbNullString)), 2)), MODE_ARGENT)

745     itmSoum.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = lColor

750     If itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = vbNullString Then
755       itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = " "
760     End If

765     Call CalculerTempsFabrication

770     Exit Sub

AfficherErreur:

775     woups "frmProjSoumElec", "AjouterNegatifDansListView", Err, Erl
End Sub

Private Sub AjouterDansListViewSoumission(ByVal dblQuantite As Double, ByVal sSousSection As String)

5       On Error GoTo AfficherErreur

10      Dim rstConfig   As ADODB.Recordset
15      Dim iIndex      As Integer
20      Dim iCompteur   As Integer
25      Dim iIDSection  As Integer
30      Dim iTagSection As Integer
35      Dim iIndexSel   As Integer
40      Dim itmSoum     As ListItem
45      Dim bSelected   As Boolean
50      Dim dblTempsMec As Double
55      Dim sDistrib    As String
60      Dim sTauxUSA    As String
65      Dim sTauxSPA    As String
70      Dim lColor      As Long
  
75      bSelected = False
  
        'S'il y a des items dans le ListView
80      If lvwSoumission.ListItems.count > 0 Then
          'Si ce n'est pas le premier qui est sélectionné
          '(le premier est sélectionné par défaut)
85        If lvwSoumission.SelectedItem.Index > 1 Then
90          bSelected = True
      
95          iIndexSel = lvwSoumission.SelectedItem.Index
100       End If
105     End If
 
        'Si le premier est sélectionné, on le considère comme si aucun n'est sélectionné
110     If bSelected = False Then
115       iIndex = TrouverIndexSection(sSousSection)
120     Else
          'Sinon, on l'ajoute à l'endroit sélectionné
125       iIndex = iIndexSel
130     End If
  
135     Set rstConfig = New ADODB.Recordset

140     Call rstConfig.Open("SELECT TauxAmericain, TauxEspagnol FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)
  
145     sTauxUSA = rstConfig.Fields("TauxAmericain")
150     sTauxSPA = rstConfig.Fields("TauxEspagnol")

155     Call rstConfig.Close
160     Set rstConfig = Nothing
  
165     Set itmSoum = lvwSoumission.ListItems.Add(iIndex)
  
170     itmSoum.Checked = True
  
        'Quantité
175     itmSoum.Text = dblQuantite
  
180     If lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_QUOTER) = "Oui" Then
185       itmSoum.Text = itmSoum.Text & "*"
190       itmSoum.ForeColor = COLOR_VERT
195       itmSoum.Bold = True
200     Else
205       itmSoum.ForeColor = COLOR_NOIR
210       itmSoum.Bold = False
215     End If
  
220     If lvwfournisseur.SelectedItem.Text = "CHOISIR ULTÉRIEUREMENT" Then
225       lColor = COLOR_MAGENTA
230     Else
235       lColor = COLOR_NOIR
240     End If

        'On met l'id de la section dans le tag du listItem
245     itmSoum.Tag = cmbSections.ItemData(cmbSections.ListIndex)
                                                                                                         
        'No d'item
250     If m_bRecherchePiece = True Then
255       itmSoum.SubItems(I_COL_SOUM_PIECE) = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM)
260     Else
265       itmSoum.SubItems(I_COL_SOUM_PIECE) = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_NO_ITEM)
270     End If

275     itmSoum.ListSubItems(I_COL_SOUM_PIECE).ForeColor = lColor
  
        'On met le nom de la sous-section dans le tag du no d'item
280     itmSoum.ListSubItems(I_COL_SOUM_PIECE).Tag = sSousSection
  
        'On met la description en francais dans la colonne et la description en anglais
        'dans le tag
285     If m_eLangage = ANGLAIS Then
290       If m_bRecherchePiece = True Then
295         itmSoum.SubItems(I_COL_SOUM_DESCR) = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_DESCR_EN)
300         itmSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_DESCR_FR)
305       Else
310         itmSoum.SubItems(I_COL_SOUM_DESCR) = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_DESCR_EN)
315         itmSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_DESCR_FR)
320       End If
325     Else
330       If m_bRecherchePiece = True Then
335         itmSoum.SubItems(I_COL_SOUM_DESCR) = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_DESCR_FR)
340         itmSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_DESCR_EN)
345       Else
350         itmSoum.SubItems(I_COL_SOUM_DESCR) = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_DESCR_FR)
355         itmSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_DESCR_EN)
360       End If
365     End If

370     itmSoum.ListSubItems(I_COL_SOUM_DESCR).ForeColor = lColor
          
        'On met le nom du fabricant dans la colonne et l'ordre de la section dans le tag
375     If m_bRecherchePiece = True Then
380       itmSoum.SubItems(I_COL_SOUM_MANUFACT) = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_MANUFACT)
385     Else
390       itmSoum.SubItems(I_COL_SOUM_MANUFACT) = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_MANUFACT)
395     End If

400     itmSoum.ListSubItems(I_COL_SOUM_MANUFACT).Tag = cmbSections.ListIndex + 1 'Ordre de la section
  
405     itmSoum.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = lColor

        'Prix listé
410     If Trim$(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_LIST)) = vbNullString Then
415       itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion("0", MODE_ARGENT, 4)
420     Else
425       If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "USA" Then
430         itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_LIST)) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
435       Else
440         If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "SPA" Then
445           itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_LIST)) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
450         Else
455           itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_LIST), MODE_ARGENT, 4)
460         End If
465       End If
470     End If

475     itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PRIX_LIST).Tag
       
480     itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = lColor
     
        'Si il y a un prix net, on le met l'escompte et le prix net sinon, on prend le prix
        'spécial pour mettre dans le prix net
485     If Trim$(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_NET)) <> vbNullString Then
490       If Trim$(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_ESCOMPTE)) <> vbNullString Then
495         itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_ESCOMPTE), MODE_POURCENT)
500       Else
505         itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion(0, MODE_POURCENT)
510       End If

515       If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "USA" Then
520         itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_NET)) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
525       Else
530         If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "SPA" Then
535           itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_NET)) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
540         Else
545           itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_NET), MODE_ARGENT, 4)
550         End If
555       End If
560     Else
565       If Trim$(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_SP)) <> vbNullString Then
570         If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "USA" Then
575           itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_SP)) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
580         Else
585           If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "SPA" Then
590             itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_SP)) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
595           Else
600             itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_SP), MODE_ARGENT, 4)
605           End If
610         End If
615       Else
620         itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion("0", MODE_POURCENT)
625         itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion("0", MODE_ARGENT, 4)
630       End If
635     End If
     
640     itmSoum.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = lColor
645     itmSoum.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = lColor
      
        'On met le fournisseur dans la colonne et l'id dans le tag
650     If lvwfournisseur.SelectedItem.Text = "CHOISIR ULTÉRIEUREMENT" Then
655       sDistrib = vbNullString
660     Else
665       sDistrib = lvwfournisseur.SelectedItem.Text
670     End If

675     itmSoum.SubItems(I_COL_SOUM_DISTRIB) = sDistrib
680     itmSoum.ListSubItems(I_COL_SOUM_DISTRIB).Tag = lvwfournisseur.SelectedItem.Tag
   
685     itmSoum.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = lColor
    
        'Temps
690     If m_bRecherchePiece = True Then
695       itmSoum.SubItems(I_COL_SOUM_TEMPS) = Replace(lvwPieceTrouve.SelectedItem.Tag, ".", ",")
700     Else
705       itmSoum.SubItems(I_COL_SOUM_TEMPS) = Replace(lvwPieces.SelectedItem.Tag, ".", ",")
710     End If

715     itmSoum.ListSubItems(I_COL_SOUM_TEMPS).ForeColor = lColor
   
        'Si le temps n'est pas vide
720     If Trim$(itmSoum.SubItems(I_COL_SOUM_TEMPS)) <> vbNullString Then
          'On calcul le temps * quantité pour la colonne montage
725       itmSoum.SubItems(I_COL_SOUM_MONTAGE) = CDbl(itmSoum.SubItems(I_COL_SOUM_TEMPS)) * CDbl(Replace(itmSoum.Text, "*", vbNullString))
730     Else
735       itmSoum.SubItems(I_COL_SOUM_MONTAGE) = vbNullString
740     End If
      
745     itmSoum.ListSubItems(I_COL_SOUM_MONTAGE).ForeColor = lColor
     
        'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
750     itmSoum.SubItems(I_COL_SOUM_TOTAL) = Conversion(Round(Replace(itmSoum.Text, "*", vbNullString) * itmSoum.SubItems(I_COL_SOUM_PRIX_NET) * CSng(m_sProfit), 2), MODE_ARGENT, 2)
      
755     itmSoum.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = lColor

760     itmSoum.ListSubItems(I_COL_SOUM_TOTAL).Tag = lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag
      
        'Pour le profit, c'est le prix total - (prix net * quantité)
765     itmSoum.SubItems(I_COL_SOUM_PROFIT) = Conversion(CStr(Round(itmSoum.SubItems(I_COL_SOUM_TOTAL) - (itmSoum.SubItems(I_COL_SOUM_PRIX_NET) * Replace(itmSoum.Text, "*", vbNullString)), 2)), MODE_ARGENT)

770     itmSoum.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = lColor

775     If itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = vbNullString Then
780       itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = " "
785     End If

790     Call CalculerTempsFabrication

795     Call itmSoum.EnsureVisible

800     Exit Sub

AfficherErreur:

805     woups "frmProjSoumElec", "AjouterDansListViewSoumission", Err, Erl
End Sub

Private Sub CalculerTempsFabrication()

5       On Error GoTo AfficherErreur

10      Dim dblTempsFab As Double
15      Dim iCompteur   As Integer

        'Pour chaque élément du listView
20      For iCompteur = 1 To lvwSoumission.ListItems.count
25        If Trim$(lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_MONTAGE)) <> vbNullString Then
            'On additionne le temps
30          dblTempsFab = dblTempsFab + CDbl(Replace(Trim$(lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_MONTAGE)), ".", ","))
35        End If
40      Next
        
45      m_sTempsFabrication = Replace(dblTempsFab / 10, ".", ",")

50      Exit Sub

AfficherErreur:

55      woups "frmProjSoumElec", "CalculerTempsFabrication", Err, Erl
End Sub

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

65      woups "frmProjSoumElec", "VerifierEmplacement", Err, Erl
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

50      woups "frmProjSoumElec", "ValeurParDefaut", Err, Erl
End Sub

Private Sub RemplirListViewProjSoum(ByVal sNoProjSoum As String)

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
55      Dim bBold         As Boolean

60      Set rstProjSoum = New ADODB.Recordset
65      Set rstSection = New ADODB.Recordset
70      Set rstFRS = New ADODB.Recordset
  
75      Call lvwSoumission.ListItems.Clear
  
80      bPremierEnr = True
  
85      If m_eType = TYPE_PROJET Then
90        Call rstProjSoum.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoProjSoum & "' ORDER BY CInt(OrdreSection), NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
95      Else
100       Call rstProjSoum.Open("SELECT * FROM GRB_Soumission_Pieces WHERE IDSoumission = '" & sNoProjSoum & "' ORDER BY CInt(OrdreSection), NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
105     End If
  
110     If m_eLangage = ANGLAIS Then
115       sSection = "NomSectionEN"
120     Else
125       sSection = "NomSectionFR"
130     End If
  
135     Do While Not rstProjSoum.EOF
140       Set itmProjSoum = lvwSoumission.ListItems.Add
          
          'Si c'est le premier enregistrement, il faut ajouter la section et la sous-section
145       If bPremierEnr = True Then
150         iOrdreSection = rstProjSoum.Fields("OrdreSection")
155         sSousSection = rstProjSoum.Fields("SousSection")
    
            'Pour avoir le nom de la section
160         Call rstSection.Open("SELECT " & sSection & " FROM GRB_SoumProjSectionElec WHERE IDSection = " & rstProjSoum.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
      
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
        
285           Call rstSection.Open("SELECT " & sSection & " FROM GRB_SoumProjSectionElec WHERE IDSection = " & rstProjSoum.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
        
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

480       If rstProjSoum.Fields("PieceExtraChargeable") = True Then
485         lColor = COLOR_BLEU
490         bBold = True
495       Else
500         If rstProjSoum.Fields("PieceExtraNonChargeable") = True Then
505           lColor = COLOR_ROSE
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
    
          'On met la quantité en vert avec un astérix si il est quoté
745       If rstProjSoum.Fields("Quoté") = True Then
750         itmProjSoum.Text = itmProjSoum.Text & "*"
755         itmProjSoum.ForeColor = COLOR_VERT
760         itmProjSoum.Bold = True
765       Else
770         itmProjSoum.ForeColor = COLOR_NOIR
775         itmProjSoum.Bold = False
780       End If

          'Facturation
785       If m_eType = TYPE_PROJET Then
790         If g_bModificationProjetsElec = True Then
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
1030          If Not IsNull(rstProjSoum.Fields("DESC_EN")) Then
1035            itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = rstProjSoum.Fields("DESC_EN")
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
1120        If Not IsNull(rstProjSoum.Fields("PRIX_LIST")) Then
1125          If rstProjSoum.Fields("PRIX_LIST") <> "" Then
1130            itmProjSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(rstProjSoum.Fields("PRIX_LIST"), MODE_ARGENT, 4)
1135          Else
1140            itmProjSoum.SubItems(I_COL_SOUM_PRIX_LIST) = " "
1145          End If
1150        Else
1155          itmProjSoum.SubItems(I_COL_SOUM_PRIX_LIST) = " "
1160        End If

1165        itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = lColor
1170        itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).Bold = bBold
       
1175        itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = rstProjSoum.Fields("PrixOrigine")

            'Escompte
1180        If Trim(rstProjSoum.Fields("Escompte")) <> vbNullString Then
1185          itmProjSoum.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion(rstProjSoum.Fields("Escompte"), MODE_POURCENT)
1190        Else
1195          itmProjSoum.SubItems(I_COL_SOUM_ESCOMPTE) = " "
1200        End If
     
1205        itmProjSoum.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = lColor
  
1210        itmProjSoum.ListSubItems(I_COL_SOUM_ESCOMPTE).Bold = bBold
      
1215        If Not IsNull(rstProjSoum.Fields("PRIX_NET")) Then
1220          If rstProjSoum.Fields("PRIX_NET") <> "" Then
1225            itmProjSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(rstProjSoum.Fields("PRIX_NET"), MODE_ARGENT, 4)
1230          Else
1235            itmProjSoum.SubItems(I_COL_SOUM_PRIX_NET) = " "
1240          End If
1245        Else
1250          itmProjSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion("", MODE_ARGENT, 4)
1255        End If
            
1260        itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = lColor
  
1265        itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_NET).Bold = bBold
  
1270        If m_eType = TYPE_PROJET Then
1275          itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_NET).Tag = rstProjSoum.Fields("DateRéception")
1280        End If
         
            'Fournisseur
1285        If Not IsNull(rstProjSoum.Fields("IDFRS")) And rstProjSoum.Fields("IDFRS") > 0 Then
1290          If itmProjSoum.SubItems(I_COL_SOUM_PIECE) <> "Texte" And itmProjSoum.SubItems(I_COL_SOUM_PIECE) <> "Text" Then
1295            Call rstFRS.Open("SELECT NomFournisseur FROM GRB_Fournisseur WHERE IDFRS = " & rstProjSoum.Fields("IDFRS"), g_connData, adOpenDynamic, adLockOptimistic)
  
                'On affiche le nom dans la colonne
1300            itmProjSoum.SubItems(I_COL_SOUM_DISTRIB) = rstFRS.Fields("NomFournisseur")
          
1305            itmProjSoum.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = lColor
1310            itmProjSoum.ListSubItems(I_COL_SOUM_DISTRIB).Bold = bBold
       
                'On affiche l'Id dans le tag
1315            itmProjSoum.ListSubItems(I_COL_SOUM_DISTRIB).Tag = rstProjSoum.Fields("IDFRS")
     
1320            Call rstFRS.Close
1325          End If
1330        Else
1335          itmProjSoum.SubItems(I_COL_SOUM_DISTRIB) = vbNullString
1340          itmProjSoum.ListSubItems(I_COL_SOUM_DISTRIB).Tag = 0
1345        End If
    
            'Temps
1350        If Not IsNull(rstProjSoum.Fields("Temps")) Then
1355          itmProjSoum.SubItems(I_COL_SOUM_TEMPS) = rstProjSoum.Fields("Temps")
1360        Else
1365          itmProjSoum.SubItems(I_COL_SOUM_TEMPS) = vbNullString
1370        End If

1375        itmProjSoum.ListSubItems(I_COL_SOUM_TEMPS).ForeColor = lColor

1380        itmProjSoum.ListSubItems(I_COL_SOUM_TEMPS).Bold = bBold
    
            'Montage
1385        If Not IsNull(rstProjSoum.Fields("Temps_total")) Then
1390          itmProjSoum.SubItems(I_COL_SOUM_MONTAGE) = rstProjSoum.Fields("Temps_total")
1395        Else
1400          itmProjSoum.SubItems(I_COL_SOUM_MONTAGE) = vbNullString
1405        End If
   
1410        itmProjSoum.ListSubItems(I_COL_SOUM_MONTAGE).ForeColor = lColor
 
1415        itmProjSoum.ListSubItems(I_COL_SOUM_MONTAGE).Bold = bBold
    
            'Prix total
1420        If Trim(rstProjSoum.Fields("Prix_total")) <> vbNullString Then
1425          If IsNumeric(rstProjSoum.Fields("Prix_Total")) Then
1430            itmProjSoum.SubItems(I_COL_SOUM_TOTAL) = Conversion(Round(rstProjSoum.Fields("Prix_total"), 2), MODE_ARGENT)
1435          Else
1440            itmProjSoum.SubItems(I_COL_SOUM_TOTAL) = " "
1445          End If
1450        Else
1455          itmProjSoum.SubItems(I_COL_SOUM_TOTAL) = " "
1460        End If
   
1465        itmProjSoum.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = lColor

1470        itmProjSoum.ListSubItems(I_COL_SOUM_TOTAL).Bold = bBold
      
1475        itmProjSoum.ListSubItems(I_COL_SOUM_TOTAL).Tag = rstProjSoum.Fields("Devise")
      
            'Profit
1480        If Trim(rstProjSoum.Fields("Profit_argent")) <> vbNullString Then
1485          itmProjSoum.SubItems(I_COL_SOUM_PROFIT) = Conversion(Round(rstProjSoum.Fields("Profit_Argent"), 2), MODE_ARGENT)
1490        Else
1495          itmProjSoum.SubItems(I_COL_SOUM_PROFIT) = " "
1500        End If

1505        If m_eType = TYPE_PROJET Then
1510          If rstProjSoum.Fields("PieceExtraChargeable") = True Or rstProjSoum.Fields("PieceExtraNonChargeable") = True Then
1515            itmProjSoum.ListSubItems(I_COL_SOUM_PROFIT).Tag = "EXTRA"
1520          End If
1525        End If
    
1530        itmProjSoum.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = lColor

1535        itmProjSoum.ListSubItems(I_COL_SOUM_PROFIT).Bold = bBold

1540        If Not IsNull(rstProjSoum.Fields("Commentaire")) Then
1545          itmProjSoum.SubItems(I_COL_SOUM_COMMENTAIRE) = rstProjSoum.Fields("Commentaire")
1550        Else
1555          itmProjSoum.SubItems(I_COL_SOUM_COMMENTAIRE) = vbNullString
1560        End If

1565        itmProjSoum.ListSubItems(I_COL_SOUM_COMMENTAIRE).ForeColor = lColor

1570        itmProjSoum.ListSubItems(I_COL_SOUM_COMMENTAIRE).Bold = bBold

1575        If m_eType = TYPE_PROJET Then
1580          If Not IsNull(rstProjSoum.Fields("ID")) Then
1585            itmProjSoum.SubItems(I_COL_SOUM_ID) = rstProjSoum.Fields("ID")
1590          Else
1595            itmProjSoum.SubItems(I_COL_SOUM_ID) = vbNullString
1600          End If

1605          itmProjSoum.ListSubItems(I_COL_SOUM_ID).ForeColor = lColor

1610          itmProjSoum.ListSubItems(I_COL_SOUM_ID).Bold = bBold

1615          If Not IsNull(rstProjSoum.Fields("DateCommande")) Then
1620            If Trim(rstProjSoum.Fields("DateCommande")) <> "" Then
1625              itmProjSoum.SubItems(I_COL_SOUM_DATE_COMMANDE) = rstProjSoum.Fields("DateCommande")
1630            Else
1635              itmProjSoum.SubItems(I_COL_SOUM_DATE_COMMANDE) = " "
1640            End If
1645          Else
1650            itmProjSoum.SubItems(I_COL_SOUM_DATE_COMMANDE) = " "
1655          End If

1660          itmProjSoum.ListSubItems(I_COL_SOUM_DATE_COMMANDE).ForeColor = lColor

1665          itmProjSoum.ListSubItems(I_COL_SOUM_DATE_COMMANDE).Bold = bBold

1670          itmProjSoum.ListSubItems(I_COL_SOUM_DATE_COMMANDE).Tag = rstProjSoum.Fields("NoRetour")

1675          If Not IsNull(rstProjSoum.Fields("DateRequise")) Then
1680            If Trim(rstProjSoum.Fields("DateRequise")) <> "" Then
1685              itmProjSoum.SubItems(I_COL_SOUM_DATE_REQUISE) = rstProjSoum.Fields("DateRequise")
1690            Else
1695              itmProjSoum.SubItems(I_COL_SOUM_DATE_REQUISE) = " "
1700            End If
1705          Else
1710            itmProjSoum.SubItems(I_COL_SOUM_DATE_REQUISE) = " "
1715          End If

1720          itmProjSoum.ListSubItems(I_COL_SOUM_DATE_REQUISE).ForeColor = lColor

1725          itmProjSoum.ListSubItems(I_COL_SOUM_DATE_REQUISE).Bold = bBold

1730          itmProjSoum.ListSubItems(I_COL_SOUM_DATE_REQUISE).Tag = rstProjSoum.Fields("DateRetour")

1735          If Not IsNull(rstProjSoum.Fields("NomCommande")) Then
1740            itmProjSoum.SubItems(I_COL_SOUM_NOM_COMMANDE) = rstProjSoum.Fields("NomCommande")
1745          Else
1750            itmProjSoum.SubItems(I_COL_SOUM_NOM_COMMANDE) = vbNullString
1755          End If

1760          itmProjSoum.ListSubItems(I_COL_SOUM_NOM_COMMANDE).ForeColor = lColor

1765          itmProjSoum.ListSubItems(I_COL_SOUM_NOM_COMMANDE).Bold = bBold

1770          If Not IsNull(rstProjSoum.Fields("NoSéquentiel")) Then
1775            itmProjSoum.SubItems(I_COL_SOUM_NO_SEQUENTIEL) = rstProjSoum.Fields("NoSéquentiel")
1780          Else
1785            itmProjSoum.SubItems(I_COL_SOUM_NO_SEQUENTIEL) = vbNullString
1790          End If

1795          itmProjSoum.ListSubItems(I_COL_SOUM_NO_SEQUENTIEL).ForeColor = lColor

1800          itmProjSoum.ListSubItems(I_COL_SOUM_NO_SEQUENTIEL).Bold = bBold

1805          If Not IsNull(rstProjSoum.Fields("Provenance")) Then
1810            If Trim(rstProjSoum.Fields("Provenance")) <> vbNullString Then
1815              itmProjSoum.SubItems(I_COL_SOUM_PROVENANCE) = Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 3) & "-" & rstProjSoum.Fields("Provenance")
1820            Else
1825              itmProjSoum.SubItems(I_COL_SOUM_PROVENANCE) = vbNullString
1830            End If
1835          Else
1840            itmProjSoum.SubItems(I_COL_SOUM_PROVENANCE) = vbNullString
1845          End If

1850          itmProjSoum.ListSubItems(I_COL_SOUM_PROVENANCE).ForeColor = lColor
1855          itmProjSoum.ListSubItems(I_COL_SOUM_PROVENANCE).Bold = bBold
1860        Else
1865          If Not IsNull(rstProjSoum.Fields("Provenance")) Then
1870            If Trim(rstProjSoum.Fields("Provenance")) <> vbNullString Then
1875              itmProjSoum.SubItems(I_COL_SOUMISSION_PROV) = Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 3) & "-" & rstProjSoum.Fields("Provenance")
1880            Else
1885              itmProjSoum.SubItems(I_COL_SOUMISSION_PROV) = vbNullString
1890            End If
1895          Else
1900            itmProjSoum.SubItems(I_COL_SOUMISSION_PROV) = vbNullString
1905          End If

1910          itmProjSoum.ListSubItems(I_COL_SOUMISSION_PROV).ForeColor = lColor
1915          itmProjSoum.ListSubItems(I_COL_SOUMISSION_PROV).Bold = bBold
1920        End If
1925      Else
            'Fournisseur
1930        If Not IsNull(rstProjSoum.Fields("IDFRS")) And rstProjSoum.Fields("IDFRS") > 0 Then
1935          If itmProjSoum.SubItems(I_COL_SOUM_SP_PIECE) <> "Texte" And itmProjSoum.SubItems(I_COL_SOUM_SP_PIECE) <> "Text" Then
1940            Call rstFRS.Open("SELECT NomFournisseur FROM GRB_Fournisseur WHERE IDFRS = " & rstProjSoum.Fields("IDFRS"), g_connData, adOpenDynamic, adLockOptimistic)
           
                'On affiche le nom dans la colonne
1945            itmProjSoum.SubItems(I_COL_SOUM_SP_DISTRIB) = rstFRS.Fields("NomFournisseur")
         
1950            itmProjSoum.ListSubItems(I_COL_SOUM_SP_DISTRIB).ForeColor = lColor

1955            itmProjSoum.ListSubItems(I_COL_SOUM_SP_DISTRIB).Bold = bBold
      
                'On affiche l'Id dans le tag
1960            itmProjSoum.ListSubItems(I_COL_SOUM_SP_DISTRIB).Tag = rstProjSoum.Fields("IDFRS")
      
1965            Call rstFRS.Close
1970          End If
1975        Else
1980          itmProjSoum.SubItems(I_COL_SOUM_SP_DISTRIB) = vbNullString
1985          itmProjSoum.ListSubItems(I_COL_SOUM_SP_DISTRIB).Tag = 0
1990        End If
  
            'Temps
1995        If Not IsNull(rstProjSoum.Fields("Temps")) Then
2000          itmProjSoum.SubItems(I_COL_SOUM_SP_TEMPS) = rstProjSoum.Fields("Temps")
2005        Else
2010          itmProjSoum.SubItems(I_COL_SOUM_SP_TEMPS) = vbNullString
2015        End If
             
2020        itmProjSoum.ListSubItems(I_COL_SOUM_SP_TEMPS).ForeColor = lColor

2025        itmProjSoum.ListSubItems(I_COL_SOUM_SP_TEMPS).Bold = bBold
    
            'Montage
2030        If Not IsNull(rstProjSoum.Fields("Temps_total")) Then
2035          itmProjSoum.SubItems(I_COL_SOUM_SP_MONTAGE) = rstProjSoum.Fields("Temps_total")
2040        Else
2045          itmProjSoum.SubItems(I_COL_SOUM_SP_MONTAGE) = vbNullString
2050        End If

2055        itmProjSoum.ListSubItems(I_COL_SOUM_SP_MONTAGE).ForeColor = lColor

2060        itmProjSoum.ListSubItems(I_COL_SOUM_SP_MONTAGE).Bold = bBold

2065        If Not IsNull(rstProjSoum.Fields("Commentaire")) Then
2070          itmProjSoum.SubItems(I_COL_SOUM_SP_COMMENTAIRE) = rstProjSoum.Fields("Commentaire")
2075        Else
2080          itmProjSoum.SubItems(I_COL_SOUM_SP_COMMENTAIRE) = vbNullString
2085        End If

2090        itmProjSoum.ListSubItems(I_COL_SOUM_SP_COMMENTAIRE).ForeColor = lColor

2095        itmProjSoum.ListSubItems(I_COL_SOUM_SP_COMMENTAIRE).Bold = bBold

2100        If m_eType = TYPE_PROJET Then
2105          If Not IsNull(rstProjSoum.Fields("ID")) Then
2110            itmProjSoum.SubItems(I_COL_SOUM_SP_ID) = rstProjSoum.Fields("ID")
2115          Else
2120            itmProjSoum.SubItems(I_COL_SOUM_SP_ID) = vbNullString
2125          End If

2130          itmProjSoum.ListSubItems(I_COL_SOUM_SP_ID).ForeColor = lColor

2135          itmProjSoum.ListSubItems(I_COL_SOUM_SP_ID).Bold = bBold

2140          If Not IsNull(rstProjSoum.Fields("DateCommande")) Then
2145            If Trim(rstProjSoum.Fields("DateCommande")) <> "" Then
2150              itmProjSoum.SubItems(I_COL_SOUM_SP_DATE_COMMANDE) = rstProjSoum.Fields("DateCommande")
2155            Else
2160              itmProjSoum.SubItems(I_COL_SOUM_SP_DATE_COMMANDE) = " "
2165            End If
2170          Else
2175            itmProjSoum.SubItems(I_COL_SOUM_SP_DATE_COMMANDE) = " "
2180          End If

2185          itmProjSoum.ListSubItems(I_COL_SOUM_SP_DATE_COMMANDE).ForeColor = lColor

2190          itmProjSoum.ListSubItems(I_COL_SOUM_SP_DATE_COMMANDE).Bold = bBold

2195          itmProjSoum.ListSubItems(I_COL_SOUM_SP_DATE_COMMANDE).Tag = rstProjSoum.Fields("DateCommande")

2200          If Not IsNull(rstProjSoum.Fields("DateRequise")) Then
2205            If Trim(rstProjSoum.Fields("DateRequise")) <> "" Then
2210              itmProjSoum.SubItems(I_COL_SOUM_SP_DATE_REQUISE) = rstProjSoum.Fields("DateRequise")
2215            Else
2220              itmProjSoum.SubItems(I_COL_SOUM_SP_DATE_REQUISE) = ""
2225            End If
2230          Else
2235            itmProjSoum.SubItems(I_COL_SOUM_SP_DATE_REQUISE) = " "
2240          End If

2245          itmProjSoum.ListSubItems(I_COL_SOUM_SP_DATE_REQUISE).ForeColor = lColor

2250          itmProjSoum.ListSubItems(I_COL_SOUM_SP_DATE_REQUISE).Bold = bBold

2255          itmProjSoum.ListSubItems(I_COL_SOUM_SP_DATE_REQUISE).Tag = rstProjSoum.Fields("DateRetour")

2260          itmProjSoum.SubItems(I_COL_SOUM_SP_NOM_COMMANDE) = rstProjSoum.Fields("NomCommande")

2265          itmProjSoum.ListSubItems(I_COL_SOUM_SP_NOM_COMMANDE).ForeColor = lColor

2270          itmProjSoum.ListSubItems(I_COL_SOUM_SP_NOM_COMMANDE).Bold = bBold

2275          itmProjSoum.SubItems(I_COL_SOUM_SP_NO_SEQUENTIEL) = rstProjSoum.Fields("NoSéquentiel")

2280          itmProjSoum.ListSubItems(I_COL_SOUM_SP_NO_SEQUENTIEL).ForeColor = lColor

2285          itmProjSoum.ListSubItems(I_COL_SOUM_SP_NO_SEQUENTIEL).Bold = bBold

2290          If Not IsNull(rstProjSoum.Fields("Provenance")) Then
2295            If Trim(rstProjSoum.Fields("Provenance")) <> "" Then
2300              itmProjSoum.SubItems(I_COL_SOUM_SP_PROVENANCE) = Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 3) & "-" & rstProjSoum.Fields("Provenance")
2305            Else
2310              itmProjSoum.SubItems(I_COL_SOUM_SP_PROVENANCE) = ""
2315            End If
2320          Else
2325            itmProjSoum.SubItems(I_COL_SOUM_SP_PROVENANCE) = ""
2330          End If

2335          itmProjSoum.ListSubItems(I_COL_SOUM_SP_PROVENANCE).ForeColor = lColor
2340          itmProjSoum.ListSubItems(I_COL_SOUM_SP_PROVENANCE).Bold = bBold
2345        Else
2350          If Not IsNull(rstProjSoum.Fields("Provenance")) Then
2355            If Trim(rstProjSoum.Fields("Provenance")) <> "" Then
2360              itmProjSoum.SubItems(I_COL_SOUMISSION_SP_PROV) = Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 3) & "-" & rstProjSoum.Fields("Provenance")
2365            Else
2370              itmProjSoum.SubItems(I_COL_SOUMISSION_SP_PROV) = ""
2375            End If
2380          Else
2385            itmProjSoum.SubItems(I_COL_SOUMISSION_SP_PROV) = ""
2390          End If

2395          itmProjSoum.ListSubItems(I_COL_SOUMISSION_SP_PROV).ForeColor = lColor
2400          itmProjSoum.ListSubItems(I_COL_SOUMISSION_SP_PROV).Bold = bBold
2405        End If
2410      End If
  
2415      Call rstProjSoum.MoveNext
  
2420      Call lvwSoumission.Refresh
2425    Loop

2430    If lvwSoumission.ListItems.count > 0 Then
2435      Call Deselect

2440      lvwSoumission.ListItems(1).Selected = True
2445    End If

2450    Call CalculerPrix
  
2455    Call rstProjSoum.Close
2460    Set rstProjSoum = Nothing

2465    Set rstFRS = Nothing
2470    Set rstSection = Nothing

2475    Exit Sub

AfficherErreur:

2480    woups "frmProjSoumElec", "RemplirListViewProjSoum", Err, Erl, sNoProjSoum)
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

55      Set rstProjSoum = New ADODB.Recordset
60      Set rstSection = New ADODB.Recordset
65      Set rstFRS = New ADODB.Recordset
  
70      Call lvwSoumission.ListItems.Clear
  
75      bPremierEnr = True
  
80      Call rstProjSoum.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoProjet & "' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
  
85      If m_eLangage = ANGLAIS Then
90        sSection = "NomSectionEN"
95      Else
100       sSection = "NomSectionFR"
105     End If
  
110     Do While Not rstProjSoum.EOF
115       Set itmProjSoum = lvwSoumission.ListItems.Add
          
          'Si c'est le premier enregistrement, il faut ajouter la section et la sous-section
120       If bPremierEnr = True Then
125         iOrdreSection = rstProjSoum.Fields("OrdreSection")
130         sSousSection = rstProjSoum.Fields("SousSection")
     
            'Pour avoir le nom de la section
135         Call rstSection.Open("SELECT " & sSection & " FROM GRB_SoumProjSectionElec WHERE IDSection = " & rstProjSoum.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
      
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
        
260           Call rstSection.Open("SELECT " & sSection & " FROM GRB_SoumProjSectionElec WHERE IDSection = " & rstProjSoum.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
        
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
    
          'On met la quantité en vert avec un astérix si il est quoté
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
 
              'On met la description en francais dans le tag de la description en anglais
695           If Not IsNull(rstProjSoum.Fields("DESC_FR")) Then
700             itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = rstProjSoum.Fields("DESC_FR")
705           Else
710             itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = vbNullString
715           End If
720         Else
              'Description en francais
725           If Not IsNull(rstProjSoum.Fields("DESC_FR")) Then
730             itmProjSoum.SubItems(I_COL_SOUM_DESCR) = rstProjSoum.Fields("DESC_FR")
735           Else
740             itmProjSoum.SubItems(I_COL_SOUM_DESCR) = vbNullString
745           End If
    
              'On met la description en anglais dans le tag de la description en francais
750           If Not IsNull(rstProjSoum.Fields("DESC_EN")) Then
755             itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = rstProjSoum.Fields("DESC_EN")
760           Else
765             itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = vbNullString
770           End If
775         End If
780       End If
   
785       itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).ForeColor = lColor

          'Fabricant
790       If Not IsNull(rstProjSoum.Fields("Manufact")) Then
795         itmProjSoum.SubItems(I_COL_SOUM_MANUFACT) = rstProjSoum.Fields("Manufact")
800       Else
805         itmProjSoum.SubItems(I_COL_SOUM_MANUFACT) = vbNullString
810       End If
   
815       itmProjSoum.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = lColor
    
          'On met l'ordre de la section dans le tag du fabricant
820       itmProjSoum.ListSubItems(I_COL_SOUM_MANUFACT).Tag = iOrdreSection
    
          'Prix listé
825       If Trim(rstProjSoum.Fields("Prix_List")) <> vbNullString Then
830         itmProjSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(rstProjSoum.Fields("Prix_list"), MODE_ARGENT, 4)
835       Else
840         itmProjSoum.SubItems(I_COL_SOUM_PRIX_LIST) = " "
845       End If
     
850       itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = lColor
      
855       itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = rstProjSoum.Fields("PrixOrigine")
      
          'Escompte
860       If Trim(rstProjSoum.Fields("Escompte")) <> vbNullString Then
865         itmProjSoum.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion(rstProjSoum.Fields("Escompte"), MODE_POURCENT)
870       Else
875         itmProjSoum.SubItems(I_COL_SOUM_ESCOMPTE) = " "
880       End If
    
885       itmProjSoum.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = lColor
    
          'Prix net
890       If Trim(rstProjSoum.Fields("Prix_net")) <> vbNullString Then
895         itmProjSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(rstProjSoum.Fields("Prix_net"), MODE_ARGENT, 4)
900       Else
905         itmProjSoum.SubItems(I_COL_SOUM_PRIX_NET) = " "
910       End If
          
915       itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = lColor
          
          'Fournisseur
920       If Not IsNull(rstProjSoum.Fields("IDFRS")) And rstProjSoum.Fields("IDFRS") > 0 Then
925         If itmProjSoum.SubItems(I_COL_SOUM_PIECE) <> "Texte" And itmProjSoum.SubItems(I_COL_SOUM_PIECE) <> "Text" Then
930           Call rstFRS.Open("SELECT NomFournisseur FROM GRB_Fournisseur WHERE IDFRS = " & rstProjSoum.Fields("IDFRS"), g_connData, adOpenDynamic, adLockOptimistic)
    
              'On affiche le nom dans la colonne
935           itmProjSoum.SubItems(I_COL_SOUM_DISTRIB) = rstFRS.Fields("NomFournisseur")
          
940           itmProjSoum.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = lColor
       
              'On affiche l'Id dans le tag
945           itmProjSoum.ListSubItems(I_COL_SOUM_DISTRIB).Tag = rstProjSoum.Fields("IDFRS")
      
950           Call rstFRS.Close
955         End If
960       Else
965         itmProjSoum.SubItems(I_COL_SOUM_DISTRIB) = vbNullString
970         itmProjSoum.ListSubItems(I_COL_SOUM_DISTRIB).Tag = 0
975       End If
    
          'Temps
980       If Not IsNull(rstProjSoum.Fields("Temps")) Then
985         itmProjSoum.SubItems(I_COL_SOUM_TEMPS) = rstProjSoum.Fields("Temps")
990       Else
995         itmProjSoum.SubItems(I_COL_SOUM_TEMPS) = vbNullString
1000      End If
    
1005      itmProjSoum.ListSubItems(I_COL_SOUM_TEMPS).ForeColor = lColor
    
          'Montage
1010      If Not IsNull(rstProjSoum.Fields("Temps_total")) Then
1015        itmProjSoum.SubItems(I_COL_SOUM_MONTAGE) = rstProjSoum.Fields("Temps_total")
1020      Else
1025        itmProjSoum.SubItems(I_COL_SOUM_MONTAGE) = vbNullString
1030      End If
    
1035      itmProjSoum.ListSubItems(I_COL_SOUM_MONTAGE).ForeColor = lColor
      
          'Prix total
1040      If Trim(rstProjSoum.Fields("Prix_total")) <> vbNullString Then
1045        itmProjSoum.SubItems(I_COL_SOUM_TOTAL) = Conversion(Round(rstProjSoum.Fields("Prix_total"), 2), MODE_ARGENT)
1050      Else
1055        itmProjSoum.SubItems(I_COL_SOUM_TOTAL) = " "
1060      End If
    
1065      itmProjSoum.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = lColor
      
          'Profit
1070      If Trim(rstProjSoum.Fields("Profit_argent")) <> vbNullString Then
1075        itmProjSoum.SubItems(I_COL_SOUM_PROFIT) = Conversion(Round(rstProjSoum.Fields("Profit_Argent"), 2), MODE_ARGENT)
1080      Else
1085        itmProjSoum.SubItems(I_COL_SOUM_PROFIT) = " "
1090      End If
    
1095      itmProjSoum.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = lColor

1100      If Not IsNull(rstProjSoum.Fields("Commentaire")) Then
1105        itmProjSoum.SubItems(I_COL_SOUM_COMMENTAIRE) = rstProjSoum.Fields("Commentaire")
1110      Else
1115        itmProjSoum.SubItems(I_COL_SOUM_COMMENTAIRE) = vbNullString
1120      End If

1125      itmProjSoum.ListSubItems(I_COL_SOUM_COMMENTAIRE).ForeColor = lColor
   
1130      Call rstProjSoum.MoveNext
  
1135      Call lvwSoumission.Refresh
1140    Loop
  
1145    Call rstProjSoum.Close
1150    Set rstProjSoum = Nothing

1155    Set rstFRS = Nothing
1160    Set rstSection = Nothing

1165    Exit Sub

AfficherErreur:

1170    woups "frmProjSoumElec", "RemplirListSoumissionProjet", Err, Erl
End Sub

Private Sub CalculerPrix()

5       On Error GoTo AfficherErreur

        'Méthode pour calculer le prix
10      Dim dblPrixPieces         As Double
15      Dim dblPrixTotal          As Double
20      Dim dblCommission         As Double
25      Dim dblTotalTemps         As Double
30      Dim dblProfit             As Double
35      Dim dblTotalManuel        As Double
40      Dim dblTotalImprevue      As Double
45      Dim dblGrandTotal         As Double
50      Dim dblTotalDessin        As Double
55      Dim dblTotalFabrication   As Double
60      Dim dblTotalAssemblage    As Double
65      Dim dblTotalProgInterface As Double
70      Dim dblTotalProgAutomate  As Double
75      Dim dblTotalProgRobot     As Double
80      Dim dblTotalVision        As Double
85      Dim dblTotalTest          As Double
90      Dim dblTotalInstallation  As Double
95      Dim dblTotalMiseService   As Double
100     Dim dblTotalFormation     As Double
105     Dim dblTotalGestion       As Double
110     Dim dblTotalShipping      As Double
115     Dim dblHebergement        As Double
120     Dim dblRepas              As Double
125     Dim dblTransport          As Double
130     Dim dblUniteMobile        As Double
135     Dim dblPrixEmballage      As Double
140     Dim dblTotalResteTemps    As Double
145     Dim bDemande              As Boolean
150     Dim iNbrePersonne         As Integer
155     Dim iCompteur             As Integer
        
        'Si ce n'est pas en mode affichage
160     If m_bModeAffichage = False Then
          'Pour chaque élément du listview
165       For iCompteur = 1 To lvwSoumission.ListItems.count
            'Si ce n'est pas une section
170         If lvwSoumission.ListItems(iCompteur).Tag <> vbNullString Then
              'Si ce n'est pas une sous-section
175           If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> vbNullString And lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> "Texte" And lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> "Text" Then
180             If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_DISTRIB) <> vbNullString Then
                  'On additionne le prix total
                  
185               If IsNumeric(lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_TOTAL)) And IsNumeric(lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PROFIT)) Then
190                 dblPrixPieces = dblPrixPieces + lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_TOTAL) - lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PROFIT)
195               Else
200                 Call MsgBox("La pièce " & lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) & " a un prix non numérique!", vbOKOnly, "Erreur")
205               End If
          
                  'On additionne le profit
210               If IsNumeric(Trim$(lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PROFIT))) = True Then
215                 dblProfit = dblProfit + lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PROFIT)
220               End If
225             Else
230               bDemande = True
235             End If
240           End If
245         End If
250       Next
    
          'Total des temps
255       dblTotalDessin = CDbl(m_sTempsDessin) * CDbl(m_sTauxDessin)

260       If m_bSansTemps = False Then
265         dblTotalFabrication = CDbl(m_sTempsFabrication) * CDbl(m_sTauxFabrication)
270       Else
275         dblTotalFabrication = 0
280       End If
  
285       dblTotalAssemblage = CDbl(m_sTempsAssemblage) * CDbl(m_sTauxAssemblage)
290       dblTotalProgInterface = CDbl(m_sTempsProgInterface) * CDbl(m_sTauxProgInterface)
295       dblTotalProgAutomate = CDbl(m_sTempsProgAutomate) * CDbl(m_sTauxProgAutomate)
300       dblTotalProgRobot = CDbl(m_sTempsProgRobot) * CDbl(m_sTauxProgRobot)
305       dblTotalVision = CDbl(m_sTempsVision) * CDbl(m_sTauxVision)
310       dblTotalTest = CDbl(m_sTempsTest) * CDbl(m_sTauxTest)
315       dblTotalInstallation = CDbl(m_sTempsInstallation) * CDbl(m_sTauxInstallation)
320       dblTotalMiseService = CDbl(m_sTempsMiseService) * CDbl(m_sTauxMiseService)
325       dblTotalFormation = CDbl(m_sTempsFormation) * CDbl(m_sTauxFormation)
330       dblTotalGestion = CDbl(m_sTempsGestion) * CDbl(m_sTauxGestion)
335       dblTotalShipping = CDbl(m_sTempsShipping) * CDbl(m_sTauxShipping)

340       dblTotalTemps = dblTotalDessin + _
                          dblTotalFabrication + _
                          dblTotalAssemblage + _
                          dblTotalProgInterface + _
                          dblTotalProgAutomate + _
                          dblTotalProgRobot + _
                          dblTotalVision + _
                          dblTotalTest + _
                          dblTotalInstallation + _
                          dblTotalMiseService + _
                          dblTotalFormation + _
                          dblTotalGestion + _
                          dblTotalShipping
            
345       If m_eType = TYPE_PROJET Then
350         dblHebergement = 0
355         dblRepas = 0
360         dblTransport = 0
365         dblUniteMobile = 0
370       Else
375         iNbrePersonne = Int(m_sNbrePersonne)
           
380         Do While iNbrePersonne > 0
385           If iNbrePersonne >= 2 Then
390             dblHebergement = dblHebergement + CDbl(m_sTempsHebergement) * CDbl(m_sTauxHebergement2)
              
395             iNbrePersonne = iNbrePersonne - 2
400           Else
405             dblHebergement = dblHebergement + CDbl(m_sTempsHebergement) * CDbl(m_sTauxHebergement1)
             
410             iNbrePersonne = iNbrePersonne - 1
415           End If
420         Loop
      
425         dblRepas = CDbl(m_sTempsRepas) * CDbl(m_sTauxRepas) * CDbl(m_sNbrePersonne)
430         dblTransport = CDbl(m_sTempsTransport) * CDbl(m_sTauxTransport)
435         dblUniteMobile = CDbl(m_sTempsUniteMobile) * CDbl(m_sTauxUniteMobile)
440       End If

445       If IsNumeric(m_sPrixEmballage) Then
450         dblPrixEmballage = CDbl(m_sPrixEmballage)
455       Else
460         dblPrixEmballage = 0
465       End If
      
470       dblTotalResteTemps = dblHebergement + dblRepas + dblTransport + dblUniteMobile + dblPrixEmballage
                                                              
475       If IsNumeric(txtPrixManuel.Text) Then
480         dblTotalManuel = CDbl(txtPrixManuel.Text)
485       Else
490         dblTotalManuel = 0
495       End If
                        
500       dblTotalImprevue = (dblPrixPieces + dblProfit) * CDbl(m_sImprevue)
    
505       dblPrixTotal = dblPrixPieces + dblProfit + dblTotalTemps + dblTotalImprevue + dblTotalManuel + dblTotalResteTemps
                        
          'Le calcul de la commission est sur les manuels (Nbre de page * prix par pages * nbre de copies) + (prix des pièces * pourcentage d'imprévus) + total des temps
510       dblCommission = dblPrixTotal * CDbl(m_sCommission)
        
          'Le prix total est le calcul des manuels (Nbre de page * prix par pages * nbre de copies) + (prix des pièces * pourcentage d'imprévus) + total des temps + total de la commission
515       dblGrandTotal = dblPrixTotal + dblCommission
                
          'Format monétaires avec 2 chiffres après la virgule
520       txtTotalPieces.Text = Conversion(CStr(Round(dblPrixPieces, 2)), MODE_ARGENT)
525       txtTotalTemps.Text = Conversion(CStr(Round(dblTotalTemps, 2)), MODE_ARGENT)
530       txtPrixTotal.Text = Conversion(CStr(Round(dblGrandTotal, 2)), MODE_ARGENT)
          
535       If bDemande = True Then
540         txtPrixTotal.ForeColor = COLOR_JAUNE
545       Else
550         txtPrixTotal.ForeColor = COLOR_ROUGE
555       End If

560       txtImprevus.Text = Conversion(CStr(Round(dblTotalImprevue, 2)), MODE_ARGENT)
565       txtCommission.Text = Conversion(CStr(Round(dblCommission, 2)), MODE_ARGENT)
570       txtProfit.Text = Conversion(CStr(Round(dblProfit, 2)), MODE_ARGENT)
575     Else
580       For iCompteur = 1 To lvwSoumission.ListItems.count
            'Si ce n'est pas une section
585         If lvwSoumission.ListItems(iCompteur).Tag <> vbNullString Then
              'Si ce n'est pas une sous-section
590           If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> vbNullString And lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> "Texte" And lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> "Text" Then
595             If m_bDroitPrix = True Then
600               If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_DISTRIB) = vbNullString Then
605                 bDemande = True
                  
610                 Exit For
615               End If
620             End If
625           End If
630         End If
635       Next

640       If bDemande = True Then
645         txtPrixTotal.ForeColor = COLOR_JAUNE
650       Else
655         txtPrixTotal.ForeColor = COLOR_ROUGE
660       End If
665     End If

670     Exit Sub

AfficherErreur:

675     woups "frmProjSoumElec", "CalculerPrix", Err, Erl
End Sub

Private Sub CalculerTempsFabricationRecordset(ByVal sNoProjSoum As String)

5       On Error GoTo AfficherErreur

10      Dim rstProjet   As ADODB.Recordset
15      Dim rstPiece    As ADODB.Recordset
20      Dim dblTempsFab As Double

        'Ouverture des tables
25      Set rstProjet = New ADODB.Recordset
30      Set rstPiece = New ADODB.Recordset

35      Call rstProjet.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)

40      Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet ='" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)

        'Pour chaque enregistrement du recordset
45      Do While Not rstPiece.EOF
          'Si le temps total n'est pas vide
50        If Trim(rstPiece.Fields("Temps_total")) <> vbNullString Then
            'On additionne le temps
55          dblTempsFab = dblTempsFab + CDbl(Replace(Trim(rstPiece.Fields("Temps_total")), ".", ","))
60        End If

65        Call rstPiece.MoveNext
70      Loop
                
75      rstProjet.Fields("TempsFabrication") = Replace(dblTempsFab / 10, ".", ",")

80      Call rstProjet.Update

85      Call rstPiece.Close
90      Set rstPiece = Nothing

95      Call rstProjet.Close
100     Set rstProjet = Nothing

105     Exit Sub

AfficherErreur:

110     woups "frmProjSoumElec", "CalculerTempsFabricationRecordset", Err, Erl
End Sub

Private Sub CalculerTotalRecordset(ByVal sNoProjSoum As String)

5       On Error GoTo AfficherErreur

        'Méthode pour calculer le prix
10      Dim rstProjSoum           As ADODB.Recordset
15      Dim rstPiece              As ADODB.Recordset
20      Dim rstPunch              As ADODB.Recordset
25      Dim dblTotalDessin        As Double
30      Dim dblTotalFabrication   As Double
35      Dim dblTotalAssemblage    As Double
40      Dim dblTotalProgInterface As Double
45      Dim dblTotalProgAutomate  As Double
50      Dim dblTotalProgRobot     As Double
55      Dim dblTotalVision        As Double
60      Dim dblTotalTest          As Double
65      Dim dblTotalInstallation  As Double
70      Dim dblTotalMiseService   As Double
75      Dim dblTotalFormation     As Double
80      Dim dblTotalGestion       As Double
85      Dim dblTotalShipping      As Double
90      Dim dblHebergement        As Double
95      Dim dblRepas              As Double
100     Dim dblTransport          As Double
105     Dim dblUniteMobile        As Double
110     Dim dblPrixEmballage      As Double
115     Dim dblTotalResteTemps    As Double
120     Dim dblPrixPieces         As Double
125     Dim dblPrixTotal          As Double
130     Dim dblCommission         As Double
135     Dim dblTotalTemps         As Double
140     Dim dblProfit             As Double
145     Dim dblTotalManuel        As Double
150     Dim dblTotalPieceImprevue As Double
155     Dim dblGrandTotal         As Double
160     Dim sDateDebut            As String
165     Dim sDateFin              As String
170     Dim sTotal                As String
175     Dim sFilterNoProjet       As String

180     Set rstProjSoum = New ADODB.Recordset

185     If m_eType = TYPE_PROJET Then
190       Call rstProjSoum.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)
195     Else
200       Call rstProjSoum.Open("SELECT * FROM GRB_SoumissionElec WHERE IDSoumission = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)
205     End If

210     If Not rstProjSoum.EOF Then
215       If m_eType = TYPE_PROJET Then
220         If Right$(sNoProjSoum, 2) = "99" Then
225           sFilterNoProjet = "LEFT(NoProjet, 6) = '" & Left$(sNoProjSoum, 6) & "'"
230         Else
235           sFilterNoProjet = "NoProjet = '" & sNoProjSoum & "'"
240         End If

245         sDateDebut = "TIMESERIAL(Left(GRB_Punch.HeureDébut,2),RIGHT(GRB_Punch.HeureDébut,2),0)"

250         sDateFin = "TIMESERIAL(Left(GRB_Punch.HeureFin,2),RIGHT(GRB_Punch.HeureFin,2),0)"

255         sTotal = "(SUM(" & sDateFin & " - " & sDateDebut & ")* 24) As Total"

260         Set rstPunch = New ADODB.Recordset

265         Call rstPunch.Open("SELECT Type, " & sTotal & " FROM GRB_Punch WHERE " & sFilterNoProjet & " And HeureFin is not null AND HeureDébut is not null GROUP BY Type", g_connData, adOpenDynamic, adLockOptimistic)

270         dblTotalDessin = 0
275         dblTotalFabrication = 0
280         dblTotalAssemblage = 0
285         dblTotalProgInterface = 0
290         dblTotalProgAutomate = 0
295         dblTotalProgRobot = 0
300         dblTotalVision = 0
305         dblTotalTest = 0
310         dblTotalInstallation = 0
315         dblTotalMiseService = 0
320         dblTotalFormation = 0
325         dblTotalGestion = 0
330         dblTotalShipping = 0

335         Do While Not rstPunch.EOF
340           If Not IsNull(rstPunch.Fields("Total")) Then
345             Select Case rstPunch.Fields("Type")
                  Case "Dessin":
350                 If Not IsNull(rstProjSoum.Fields("TauxDessin")) Then
355                   dblTotalDessin = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxDessin"))
360                 Else
365                   dblTotalDessin = 0
370                 End If

375               Case "Fabrication":
380                 If Not IsNull(rstProjSoum.Fields("TauxFabrication")) Then
385                   dblTotalFabrication = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxFabrication"))
390                 Else
395                   dblTotalFabrication = 0
400                 End If
                    
405               Case "Assemblage":
410                 If Not IsNull(rstProjSoum.Fields("TauxAssemblage")) Then
415                   dblTotalAssemblage = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxAssemblage"))
420                 Else
425                   dblTotalAssemblage = 0
430                 End If
                    
435               Case "ProgInterface":
440                 If Not IsNull(rstProjSoum.Fields("TauxProgInterface")) Then
445                   dblTotalProgInterface = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxProgInterface"))
450                 Else
455                   dblTotalProgInterface = 0
460                 End If
                    
465               Case "ProgAutomate":
470                 If Not IsNull(rstProjSoum.Fields("TauxProgAutomate")) Then
475                   dblTotalProgAutomate = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxProgAutomate"))
480                 Else
485                   dblTotalProgAutomate = 0
490                 End If
                    
495               Case "ProgRobot":
500                 If Not IsNull(rstProjSoum.Fields("TauxProgRobot")) Then
505                   dblTotalProgRobot = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxProgRobot"))
510                 Else
515                   dblTotalProgRobot = 0
520                 End If
                    
525               Case "Vision":
530                 If Not IsNull(rstProjSoum.Fields("TauxVision")) Then
535                   dblTotalVision = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxVision"))
540                 Else
545                   dblTotalVision = 0
550                 End If
                    
555               Case "Test":
560                 If Not IsNull(rstProjSoum.Fields("TauxTest")) Then
565                   dblTotalTest = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxTest"))
570                 Else
575                   dblTotalTest = 0
580                 End If
                    
585               Case "Installation":
590                 If Not IsNull(rstProjSoum.Fields("TauxInstallation")) Then
595                   dblTotalInstallation = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxInstallation"))
600                 Else
605                   dblTotalInstallation = 0
610                 End If
                    
615               Case "MiseService":
620                 If Not IsNull(rstProjSoum.Fields("TauxMiseService")) Then
625                   dblTotalMiseService = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxMiseService"))
630                 Else
635                   dblTotalMiseService = 0
640                 End If
                    
645               Case "Formation":
650                 If Not IsNull(rstProjSoum.Fields("TauxFormation")) Then
655                   dblTotalFormation = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxFormation"))
660                 Else
665                   dblTotalFormation = 0
670                 End If
                    
675               Case "Gestion":
680                 If Not IsNull(rstProjSoum.Fields("TauxGestion")) Then
685                   dblTotalGestion = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxGestion"))
690                 Else
695                   dblTotalGestion = 0
700                 End If
                    
705               Case "Shipping":
710                 If Not IsNull(rstProjSoum.Fields("TauxShipping")) Then
715                   dblTotalShipping = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxShipping"))
720                 Else
725                   dblTotalShipping = 0
730                 End If
735             End Select
740           End If
              
745           Call rstPunch.MoveNext
750         Loop

755         Call rstPunch.Close
760         Set rstPunch = Nothing
765       Else
770         If Not IsNull(rstProjSoum.Fields("TempsDessin")) And Not IsNull(rstProjSoum.Fields("TauxDessin")) Then
775           dblTotalDessin = CDbl(rstProjSoum.Fields("TempsDessin")) * CDbl(rstProjSoum.Fields("TauxDessin"))
780         Else
785           dblTotalDessin = 0
790         End If

795         If rstProjSoum.Fields("SansTemps") = False Then
800           If Not IsNull(rstProjSoum.Fields("TempsFabrication")) And Not IsNull(rstProjSoum.Fields("TauxFabrication")) Then
805             dblTotalFabrication = CDbl(rstProjSoum.Fields("TempsFabrication")) * CDbl(rstProjSoum.Fields("TauxFabrication"))
810           Else
815             dblTotalFabrication = 0
820           End If
825         Else
830           dblTotalFabrication = 0
835         End If

840         If Not IsNull(rstProjSoum.Fields("TempsAssemblage")) And Not IsNull(rstProjSoum.Fields("TauxAssemblage")) Then
845           dblTotalAssemblage = CDbl(rstProjSoum.Fields("TempsAssemblage")) * CDbl(rstProjSoum.Fields("TauxAssemblage"))
850         Else
855           dblTotalAssemblage = 0
860         End If

865         If Not IsNull(rstProjSoum.Fields("TempsProgInterface")) And Not IsNull(rstProjSoum.Fields("TauxProgInterface")) Then
870           dblTotalProgInterface = CDbl(rstProjSoum.Fields("TempsProgInterface")) * CDbl(rstProjSoum.Fields("TauxProgInterface"))
875         Else
880           dblTotalProgInterface = 0
885         End If

890         If Not IsNull(rstProjSoum.Fields("TempsProgAutomate")) And Not IsNull(rstProjSoum.Fields("TauxProgAutomate")) Then
895           dblTotalProgAutomate = CDbl(rstProjSoum.Fields("TempsProgAutomate")) * CDbl(rstProjSoum.Fields("TauxProgAutomate"))
900         Else
905           dblTotalProgAutomate = 0
910         End If

915         If Not IsNull(rstProjSoum.Fields("TempsProgRobot")) And Not IsNull(rstProjSoum.Fields("TauxProgRobot")) Then
920           dblTotalProgRobot = CDbl(rstProjSoum.Fields("TempsProgRobot")) * CDbl(rstProjSoum.Fields("TauxProgRobot"))
925         Else
930           dblTotalProgRobot = 0
935         End If

940         If Not IsNull(rstProjSoum.Fields("TempsVision")) And Not IsNull(rstProjSoum.Fields("TauxVision")) Then
945           dblTotalVision = CDbl(rstProjSoum.Fields("TempsVision")) * CDbl(rstProjSoum.Fields("TauxVision"))
950         Else
955           dblTotalVision = 0
960         End If

965         If Not IsNull(rstProjSoum.Fields("TempsTest")) And Not IsNull(rstProjSoum.Fields("TauxTest")) Then
970           dblTotalTest = CDbl(rstProjSoum.Fields("TempsTest")) * CDbl(rstProjSoum.Fields("TauxTest"))
975         Else
980           dblTotalTest = 0
985         End If

990         If Not IsNull(rstProjSoum.Fields("TempsInstallation")) And Not IsNull(rstProjSoum.Fields("TauxInstallation")) Then
995           dblTotalInstallation = CDbl(rstProjSoum.Fields("TempsInstallation")) * CDbl(rstProjSoum.Fields("TauxInstallation"))
1000        Else
1005          dblTotalInstallation = 0
1010        End If

1015        If Not IsNull(rstProjSoum.Fields("TempsMiseService")) And Not IsNull(rstProjSoum.Fields("TauxMiseService")) Then
1020          dblTotalMiseService = CDbl(rstProjSoum.Fields("TempsMiseService")) * CDbl(rstProjSoum.Fields("TauxMiseService"))
1025        Else
1030          dblTotalMiseService = 0
1035        End If

1040        If Not IsNull(rstProjSoum.Fields("TempsFormation")) And Not IsNull(rstProjSoum.Fields("TauxFormation")) Then
1045          dblTotalFormation = CDbl(rstProjSoum.Fields("TempsFormation")) * CDbl(rstProjSoum.Fields("TauxFormation"))
1050        Else
1055          dblTotalFormation = 0
1060        End If
  
1065        If Not IsNull(rstProjSoum.Fields("TempsGestion")) And Not IsNull(rstProjSoum.Fields("TauxGestion")) Then
1070          dblTotalGestion = CDbl(rstProjSoum.Fields("TempsGestion")) * CDbl(rstProjSoum.Fields("TauxGestion"))
1075        Else
1080          dblTotalGestion = 0
1085        End If

1090        If Not IsNull(rstProjSoum.Fields("TempsShipping")) And Not IsNull(rstProjSoum.Fields("TauxShipping")) Then
1095          dblTotalShipping = CDbl(rstProjSoum.Fields("TempsShipping")) * CDbl(rstProjSoum.Fields("TauxShipping"))
1100        Else
1105          dblTotalShipping = 0
1110        End If
1115      End If

1120      dblTotalTemps = dblTotalDessin + _
                          dblTotalFabrication + _
                          dblTotalAssemblage + _
                          dblTotalProgInterface + _
                          dblTotalProgAutomate + _
                          dblTotalProgRobot + _
                          dblTotalVision + _
                          dblTotalTest + _
                          dblTotalInstallation + _
                          dblTotalMiseService + _
                          dblTotalFormation + _
                          dblTotalGestion + _
                          dblTotalShipping

1125      Set rstPiece = New ADODB.Recordset

1130      If m_eType = TYPE_PROJET Then
1135        Call rstPiece.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoProjSoum & "' AND Type = 'E'", g_connData, adOpenDynamic, adLockOptimistic)
1140      Else
1145        Call rstPiece.Open("SELECT * FROM GRB_Soumission_Pieces WHERE IDSoumission = '" & sNoProjSoum & "' AND Type = 'E'", g_connData, adOpenDynamic, adLockOptimistic)
1150      End If

          'Pour chaque élément du recordset
1155      Do While Not rstPiece.EOF
1160        If Trim(rstPiece.Fields("Prix_total")) <> vbNullString Then
              'On additionne le prix total
1165          dblPrixPieces = dblPrixPieces + CDbl(rstPiece.Fields("Prix_total")) - CDbl(rstPiece.Fields("Profit_Argent"))
               
              'On additionne le profit
1170          dblProfit = dblProfit + CDbl(rstPiece.Fields("Profit_Argent"))
1175        End If

1180        Call rstPiece.MoveNext
1185      Loop

1190      Call rstPiece.Close
1195      Set rstPiece = Nothing

1200      If m_eType = TYPE_PROJET Then
1205        dblHebergement = 0
1210        dblRepas = 0
1215        dblTransport = 0
1220        dblUniteMobile = 0
1225      Else
1230        If Not IsNull(rstProjSoum.Fields("TotalHebergement")) Then
1235          dblHebergement = CDbl(rstProjSoum.Fields("TotalHebergement"))
1240        Else
1245          dblHebergement = 0
1250        End If

1255        If Not IsNull(rstProjSoum.Fields("TotalRepas")) Then
1260          dblRepas = CDbl(rstProjSoum.Fields("TotalRepas"))
1265        Else
1270          dblRepas = 0
1275        End If
     
1280        If Not IsNull(rstProjSoum.Fields("TempsTransport")) Then
1285          dblTransport = CDbl(rstProjSoum.Fields("TempsTransport")) * CDbl(rstProjSoum.Fields("TauxTransport"))
1290        Else
1295          dblTransport = 0
1300        End If

1305        If Not IsNull(rstProjSoum.Fields("TempsUniteMobile")) Then
1310          dblUniteMobile = CDbl(rstProjSoum.Fields("TempsUniteMobile")) * CDbl(rstProjSoum.Fields("TauxUniteMobile"))
1315        Else
1320          dblUniteMobile = 0
1325        End If
1330      End If

1335      If IsNumeric(rstProjSoum.Fields("PrixEmballage")) Then
1340        dblPrixEmballage = CDbl(rstProjSoum.Fields("PrixEmballage"))
1345      Else
1350        dblPrixEmballage = 0
1355      End If
         
1360      dblTotalResteTemps = dblHebergement + dblRepas + dblTransport + dblUniteMobile + dblPrixEmballage

1365      If IsNumeric(rstProjSoum.Fields("total_manuel")) Then
1370        dblTotalManuel = CDbl(rstProjSoum.Fields("total_manuel"))
1375      Else
1380        dblTotalManuel = 0
1385      End If

1390      dblTotalPieceImprevue = (dblPrixPieces + dblProfit) * (1 + CDbl(rstProjSoum.Fields("Imprevue")))

1395      dblPrixTotal = dblTotalTemps + dblTotalManuel + dblTotalPieceImprevue + dblTotalResteTemps

          'Le calcul de la commission est sur les manuels (Nbre de page * prix par pages * nbre de copies) + (prix des pièces * pourcentage d'imprévus) + total des temps
1400      dblCommission = dblPrixTotal * CDbl(rstProjSoum.Fields("Commission"))

          'Le prix total est le calcul des manuels (Nbre de page * prix par pages * nbre de copies) + (prix des pièces * pourcentage d'imprévus) + total des temps + total de la commission
1405      dblGrandTotal = dblPrixTotal + dblCommission

          'Format monétaire avec 2 chiffres après la virgule
1410      rstProjSoum.Fields("total_commission") = dblCommission
1415      rstProjSoum.Fields("Total_manuel") = dblTotalManuel
1420      rstProjSoum.Fields("Total_temps") = dblTotalTemps
1425      rstProjSoum.Fields("total_imprevue") = dblTotalPieceImprevue - (dblPrixPieces + dblProfit)
1430      rstProjSoum.Fields("total_piece") = dblPrixPieces
1435      rstProjSoum.Fields("Total_Commission") = Conversion(CStr(Round(dblCommission, 2)), MODE_ARGENT)
1440      rstProjSoum.Fields("PrixTotal") = Conversion(CStr(Round(dblGrandTotal, 2)), MODE_ARGENT)
1445      rstProjSoum.Fields("Total_Profit") = Conversion(CStr(Round(dblProfit, 2)), MODE_ARGENT)

1450      Call rstProjSoum.Update
1455    Else
1460      If m_eType = TYPE_PROJET Then
1465        Call MsgBox("Le projet " & sNoProjSoum & " est inexistant!", vbOKOnly, "Erreur")
1470      Else
1475        Call MsgBox("La soumission " & sNoProjSoum & " est inexistant!", vbOKOnly, "Erreur")
1480      End If
1485    End If

1490    Call rstProjSoum.Close
1495    Set rstProjSoum = Nothing

1500    Exit Sub

AfficherErreur:

1505    woups "frmProjSoumElec", "CalculerTotalRecordset", Err, Erl, sNoProjSoum)
End Sub

Private Sub CalculerPrixFacturation(ByVal sNoFacturation As String, ByRef sCommission As String, ByRef sPrixTotal As String, ByRef sProfit As String, ByRef sTempsFabrication As String, ByRef sTotalPiece As String, ByRef sImprevue As String, ByRef sTotalTemps As String, ByRef sManuel As String)

5       On Error GoTo AfficherErreur

        'Méthode pour calculer le prix
10      Dim iCompteur             As Integer
15      Dim dblTotalDessin        As Double
20      Dim dblTotalFabrication   As Double
25      Dim dblTotalAssemblage    As Double
30      Dim dblTotalProgInterface As Double
35      Dim dblTotalProgAutomate  As Double
40      Dim dblTotalProgRobot     As Double
45      Dim dblTotalVision        As Double
50      Dim dblTotalTest          As Double
55      Dim dblTotalInstallation  As Double
60      Dim dblTotalMiseService   As Double
65      Dim dblTotalFormation     As Double
70      Dim dblTotalGestion       As Double
75      Dim dblTotalShipping      As Double
80      Dim dblPrixPieces         As Double
85      Dim dblPrixTotal          As Double
90      Dim dblCommission         As Double
95      Dim dblTotalTemps         As Double
100     Dim dblProfit             As Double
105     Dim dblTotalManuel        As Double
110     Dim dblTotalPieceImprevue As Double
115     Dim dblGrandTotal         As Double
120     Dim dblTempsFabrication   As Double
    
125     dblTotalDessin = CDbl(m_sTempsDessin) * CDbl(m_sTauxDessin)

130     If m_bSansTemps = False Then
135       dblTotalFabrication = CDbl(m_sTempsFabrication) * CDbl(m_sTauxFabrication)
140     Else
145       dblTotalFabrication = 0
150     End If

155     dblTotalAssemblage = CDbl(m_sTempsAssemblage) * CDbl(m_sTauxAssemblage)
160     dblTotalProgInterface = CDbl(m_sTempsProgInterface) * CDbl(m_sTauxProgInterface)
165     dblTotalProgAutomate = CDbl(m_sTempsProgAutomate) * CDbl(m_sTauxProgAutomate)
170     dblTotalProgRobot = CDbl(m_sTempsProgRobot) * CDbl(m_sTauxProgRobot)
175     dblTotalVision = CDbl(m_sTempsVision) * CDbl(m_sTauxVision)
180     dblTotalTest = CDbl(m_sTempsTest) * CDbl(m_sTauxTest)
185     dblTotalInstallation = CDbl(m_sTempsInstallation) * CDbl(m_sTauxInstallation)
190     dblTotalMiseService = CDbl(m_sTempsMiseService) * CDbl(m_sTauxMiseService)
195     dblTotalFormation = CDbl(m_sTempsFormation) * CDbl(m_sTauxFormation)
200     dblTotalGestion = CDbl(m_sTempsGestion) * CDbl(m_sTauxGestion)
205     dblTotalShipping = CDbl(m_sTempsShipping) * CDbl(m_sTauxShipping)
    
210     dblTotalTemps = dblTotalDessin + _
                        dblTotalFabrication + _
                        dblTotalAssemblage + _
                        dblTotalProgInterface + _
                        dblTotalProgAutomate + _
                        dblTotalProgRobot + _
                        dblTotalVision + _
                        dblTotalTest + _
                        dblTotalInstallation + _
                        dblTotalMiseService + _
                        dblTotalFormation + _
                        dblTotalGestion + _
                        dblTotalShipping
    
  
        'Pour chaque élément du listview
215     For iCompteur = 1 To lvwSoumission.ListItems.count
          'Si ce n'est pas une section
220       If lvwSoumission.ListItems(iCompteur).Tag <> vbNullString Then
            'Si ce n'est pas une sous-section
225         If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> vbNullString And lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> "Texte" And lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> "Text" Then
230           If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_FACTURATION) = sNoFacturation Then
235             If Trim$(lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_TOTAL)) <> vbNullString Then
                  'On additionne le prix total
240               dblPrixPieces = dblPrixPieces + lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_TOTAL)
         
                  'On additionne le profit
245               dblProfit = dblProfit + lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PROFIT)

                  'Calcul des heures de fabrication
250               If m_bSansTemps = False Then
255                 If Trim$(lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_MONTAGE)) <> vbNullString Then
260                   dblTempsFabrication = dblTempsFabrication + lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_MONTAGE)
265                 End If
270               End If
275             End If
280           End If
285         End If
290       End If
295     Next
                    
300     If IsNumeric(txtPrixManuel.Text) Then
305       dblTotalManuel = CDbl(txtPrixManuel.Text)
310     Else
315       dblTotalManuel = 0
320     End If
                       
325     dblTotalPieceImprevue = dblPrixPieces * (1 + CDbl(m_sImprevue))
    
330     dblPrixTotal = dblTotalTemps + dblTotalManuel + dblTotalPieceImprevue
                       
        'Le calcul de la commission est sur les manuels (Nbre de page * prix par pages * nbre de copies) + (prix des pièces * pourcentage d'imprévus) + total des temps
335     dblCommission = dblPrixTotal * CDbl(m_sCommission)
        
        'Le prix total est le calcul des manuels (Nbre de page * prix par pages * nbre de copies) + (prix des pièces * pourcentage d'imprévus) + total des temps + total de la commission
340     dblGrandTotal = dblPrixTotal + dblCommission
                
        'Format monétaires avec 2 chiffres après la virgule
345     sCommission = Conversion(CStr(Round(dblCommission, 2)), MODE_ARGENT)
350     sPrixTotal = Conversion(CStr(Round(dblGrandTotal, 2)), MODE_ARGENT)
355     sProfit = Conversion(CStr(Round(dblProfit, 2)), MODE_ARGENT)
360     sTempsFabrication = dblTempsFabrication
365     sImprevue = Conversion(CStr(Round(dblPrixPieces * CDbl(m_sImprevue), 2)), MODE_ARGENT)
370     sManuel = Conversion(CStr(Round(dblTotalManuel, 2)), MODE_ARGENT)
375     sTotalPiece = Conversion(CStr(Round(dblPrixPieces, 2)), MODE_ARGENT)
380     sTotalTemps = Conversion(CStr(Round(dblTotalTemps, 2)), MODE_ARGENT)

385     Exit Sub

AfficherErreur:

390     woups "frmProjSoumElec", "CalculerPrix", Err, Erl
End Sub

Private Sub ChoisirFournisseur()

5       On Error GoTo AfficherErreur

        'On ajoute la pièce dans lvwSoumission
10      Dim sQuantite    As String
15      Dim sSousSection As String
20      Dim bDemanderSS  As Boolean
25      Dim sParams      As String
      
        'Si l'utilisateur a déjà choisi un emplacement, il ne faut pas
        'lui demander dans quelle sous-section
        
        'Si il y a des enregistrements dans le listview
30      If lvwSoumission.ListItems.count > 0 Then
          'Si le premier n'est pas sélectionné.. celui-ci est sélectionné par défaut
35        If lvwSoumission.SelectedItem.Index > 1 Then
            'Si l'emplacement est valide
40          If VerifierEmplacement(lvwSoumission.SelectedItem.Index) = True Then
              'Si c'est une sous-section
45            If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).Tag = vbNullString Then
                'Si l'autre d'au dessus est une section
50              If lvwSoumission.ListItems(lvwSoumission.SelectedItem.Index - 1).Tag = vbNullString Then
                  'Message d'erreur
55                Call MsgBox("Vous ne pouvez pas mettre une pièce entre une section et une sous-section", vbOKOnly, "Erreur")
          
60                frafournisseur.Visible = False
          
                  'Il faut resélectionné le premier pour faire comme s'il n'était plus
                  'sélectionné
65                Call Deselect
                  
70                lvwSoumission.ListItems(1).Selected = True
          
75                Exit Sub
80              Else
                  'Sinon, on prend le tag de la section d'en haut
85                sSousSection = lvwSoumission.ListItems(lvwSoumission.SelectedItem.Index - 1).ListSubItems(I_COL_SOUM_PIECE).Tag
90              End If
95            Else
                'On prend le tag de l'élément sélectionné
100             sSousSection = lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).Tag
105           End If
110         Else
115           If lvwSoumission.ListItems(lvwSoumission.SelectedItem.Index - 1).ListSubItems(I_COL_SOUM_MANUFACT).Tag <> "" Then
120             If MsgBox("Vous essayez d'ajouter une pièce de la section " & cmbSections.Text & " dans la section " & cmbSections.LIST(lvwSoumission.ListItems(lvwSoumission.SelectedItem.Index - 1).ListSubItems(I_COL_SOUM_MANUFACT).Tag - 1) & vbNewLine & "Voulez-vous ajouter la pièce dans la section " & cmbSections.LIST(lvwSoumission.ListItems(lvwSoumission.SelectedItem.Index - 1).ListSubItems(I_COL_SOUM_MANUFACT).Tag - 1), vbYesNo, "Erreur") = vbYes Then
125               cmbSections.ListIndex = lvwSoumission.ListItems(lvwSoumission.SelectedItem.Index - 1).ListSubItems(I_COL_SOUM_MANUFACT).Tag - 1

130               Call ChoisirFournisseur
135             End If
        
140             frafournisseur.Visible = False
        
                'Il faut resélectionné le premier pour faire comme si il n'était plus
                'sélectionné
145             Call Deselect
              
150             lvwSoumission.ListItems(1).Selected = True
        
155             Exit Sub
160           Else
165             Call MsgBox("Impossible d'ajouter entre une section et une sous-section!", vbOKOnly, "Erreur")

170             Exit Sub
175           End If
180         End If
185       Else
190         bDemanderSS = True
195       End If
200     Else
          'Sinon, on demande la section
205       bDemanderSS = True
210     End If
  
        'Saisie de la quantité
215     sQuantite = InputBox("Quelle est la quantité?")

220     sQuantite = Replace(sQuantite, ".", ",")
    
225     If sQuantite <> vbNullString Then
230       If Not IsNumeric(sQuantite) Then
235         Call MsgBox("Quantité non numérique!", vbOKOnly, "Erreur")
      
240         Exit Sub
245       Else
250         If sQuantite < 0 Then
255           If lvwfournisseur.SelectedItem.Text = "CHOISIR ULTÉRIEUREMENT" Then
260             Call MsgBox("Impossible de faire une demande de prix sur une pièce négative!", vbOKOnly, "Erreur")

265             Exit Sub
270           End If
275         End If
280       End If
285     Else
290       Exit Sub
295     End If

300     If bDemanderSS = True Then
305       If m_sSousSection <> S_PAS_SOUS_SECTION Then
310         sSousSection = InputBox("Quelle est la sous-section?", , m_sSousSection)
315       Else
320         sSousSection = InputBox("Quelle est la sous-section?")
325       End If
330     End If
    
        'Si la sous-section est vide
335     If sSousSection = vbNullString Then
          'On initialise la sous-section à "PAS DE SOUS-SECTIONS"
340       sSousSection = S_PAS_SOUS_SECTION
345       m_sSousSection = vbNullString
350     Else
355       m_sSousSection = sSousSection
360     End If

365     If sQuantite < 0 Then
370       If m_eType = TYPE_PROJET Then
375         If Right$(txtNoProjSoum.Text, 2) >= 60 And Right$(txtNoProjSoum.Text, 2) <= 98 Then
380           Call AjouterNegatifDansListView(CDbl(sQuantite), sSousSection)
385         Else
390           Call AjouterDansListViewSoumission(CDbl(sQuantite), sSousSection)
395         End If
400       Else
405         Call AjouterDansListViewSoumission(CDbl(sQuantite), sSousSection)
410       End If
415     Else
420       Call AjouterDansListViewSoumission(CDbl(sQuantite), sSousSection)
425     End If
  
        'Calcul des prix
430     Call CalculerPrix
  
        'On cache le listview
435     frafournisseur.Visible = False
  
        'Resélectionne le premier élément du listview
440     If lvwSoumission.ListItems.count > 0 Then
445       Call Deselect

450       lvwSoumission.ListItems(1).Selected = True
455     End If

460     Exit Sub

AfficherErreur:

465     If Err.number = 13 And Erl = 110 Then
470       sParams = "cmbSections.Text : " & cmbSections.Text & "   " & _
                    "No Proj/Soum : " & txtNoProjSoum.Text & "   " & _
                    "lvwSoumission.SelectedItem.Index - 1 : " & lvwSoumission.SelectedItem.Index - 1 & "   " & _
                    "lvwSoumission.ListItems(lvwSoumission.SelectedItem.Index - 1).ListSubItems(I_COL_SOUM_MANUFACT).Tag : " & lvwSoumission.ListItems(lvwSoumission.SelectedItem.Index - 1).ListSubItems(I_COL_SOUM_MANUFACT).Tag
            
475       woups "frmProjSoumElec", "ChoisirFournisseur", Err, Erl, sParams)
480     Else
485       woups "frmProjSoumElec", "ChoisirFournisseur", Err, Erl
490     End If
End Sub

Private Sub ChoisirFournisseurMateriel()

5       On Error GoTo AfficherErreur

        'On ajoute la pièce en négatif dans le ListView
10      Dim rstProjet  As ADODB.Recordset
15      Dim rstConfig  As ADODB.Recordset
20      Dim itmAncien  As ListItem
25      Dim itmNouveau As ListItem
30      Dim sQuantite  As String
35      Dim sExtra     As String
40      Dim sTauxUSA   As String
45      Dim sTauxSPA   As String

50      If m_bChangementFRS = True Then
55        If lvwfournisseur.SelectedItem.Text <> "CHOISIR ULTÉRIEUREMENT" Then
60          Set rstConfig = New ADODB.Recordset

65          Call rstConfig.Open("SELECT TauxAmericain, TauxEspagnol FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

70          sTauxUSA = rstConfig.Fields("TauxAmericain")
75          sTauxSPA = rstConfig.Fields("TauxEspagnol")

80          Call rstConfig.Close
85          Set rstConfig = Nothing

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

405           If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_ID) <> "" Then
410             lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_ID).ForeColor = COLOR_MAGENTA
415           End If
420         End If

425         lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_DESCR).ForeColor = COLOR_MAGENTA
430         lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = COLOR_MAGENTA
435         lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = COLOR_MAGENTA
440         lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_MONTAGE).ForeColor = COLOR_MAGENTA
445         lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_MAGENTA
450         lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = COLOR_MAGENTA
455         lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = COLOR_MAGENTA
460         lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = COLOR_MAGENTA
465         lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_TEMPS).ForeColor = COLOR_MAGENTA
470         lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = COLOR_MAGENTA

475         Call lvwSoumission.Refresh
480       End If

485       Call CalculerPrix
            
          'On cache le listview
490       frafournisseur.Visible = False

495       m_bPieceInutile = False
500       m_bChangementFRS = False
505     Else
510       If Right$(txtNoProjSoum.Text, 2) >= "00" And Right$(txtNoProjSoum.Text, 2) <= "19" Then
515         sExtra = InputBox("Dans quel extra le retour doit être fait ? (2 chiffres seulement)")

520         If Len(sExtra) <> 2 Then
525           Call MsgBox("Format incorrect!", vbOKOnly, "Erreur")

530           Exit Sub
535         End If

540         If Not IsNumeric(sExtra) Then
545           Call MsgBox("L'extra doit être numérique!", vbOKOnly, "Erreur")

550           Exit Sub
555         End If

560         If sExtra < 60 Or sExtra > 98 Then
565           Call MsgBox("L'extra doit être entre 60 et 98!", vbOKOnly, "Erreur")

570           Exit Sub
575         End If

580         Set rstProjet = New ADODB.Recordset

585         Call rstProjet.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sExtra & "'", g_connData, adOpenDynamic, adLockOptimistic)

590         If rstProjet.EOF Then
595           Call MsgBox("Le projet " & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sExtra & " n'existe pas!", vbOKOnly, "Erreur")

600           Call rstProjet.Close
605           Set rstProjet = Nothing

610           Exit Sub
615         Else
620           Call rstProjet.Close
625           Set rstProjet = Nothing
630         End If
635       End If

          'Saisie de la quantité
640       sQuantite = InputBox("Quelle est la quantité?")

645       sQuantite = Replace(sQuantite, ".", ",")

650       sQuantite = Replace(sQuantite, "-", "")

655       If sQuantite <> vbNullString Then
660         If Not IsNumeric(sQuantite) Then
665           Call MsgBox("Quantité non numérique!", vbOKOnly, "Erreur")

670           Exit Sub
675         End If
680       Else
685         Exit Sub
690       End If

695       If CDbl(sQuantite) <= CDbl(Replace(lvwSoumission.SelectedItem.Text, "*", vbNullString)) Then
700         Set itmAncien = lvwSoumission.SelectedItem
705         Set itmNouveau = lvwSoumission.ListItems.Add(itmAncien.Index + 1)

710         itmNouveau.Checked = itmAncien.Checked

            'Quantité
715         itmNouveau.Text = "-" & sQuantite

            'On met l'id de la section dans le tag du listItem
720         itmNouveau.Tag = itmAncien.Tag

            'No d'item
725         itmNouveau.SubItems(I_COL_SOUM_PIECE) = itmAncien.SubItems(I_COL_SOUM_PIECE)

            'On met le nom de la sous-section dans le tag du no d'item
730         itmNouveau.ListSubItems(I_COL_SOUM_PIECE).Tag = itmAncien.ListSubItems(I_COL_SOUM_PIECE).Tag

            'On met la description en francais dans la colonne et la description en anglais
            'dans le tag
735         itmNouveau.SubItems(I_COL_SOUM_DESCR) = itmAncien.SubItems(I_COL_SOUM_DESCR)
740         itmNouveau.ListSubItems(I_COL_SOUM_DESCR).Tag = itmAncien.ListSubItems(I_COL_SOUM_DESCR).Tag

            'On met le nom du fabricant dans la colonne et l'ordre de la section dans le tag
745         itmNouveau.SubItems(I_COL_SOUM_MANUFACT) = itmAncien.SubItems(I_COL_SOUM_MANUFACT)
750         itmNouveau.ListSubItems(I_COL_SOUM_MANUFACT).Tag = itmAncien.ListSubItems(I_COL_SOUM_MANUFACT).Tag
            
            'Prix listé
755         itmNouveau.SubItems(I_COL_SOUM_PRIX_LIST) = itmAncien.SubItems(I_COL_SOUM_PRIX_LIST)

760         itmNouveau.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = itmAncien.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag

765         itmNouveau.SubItems(I_COL_SOUM_ESCOMPTE) = itmAncien.SubItems(I_COL_SOUM_ESCOMPTE)

770         itmNouveau.SubItems(I_COL_SOUM_PRIX_NET) = itmAncien.SubItems(I_COL_SOUM_PRIX_NET)

            'On met le fournisseur dans la colonne et l'id dans le tag
775         itmNouveau.SubItems(I_COL_SOUM_DISTRIB) = lvwfournisseur.SelectedItem.Text
780         itmNouveau.ListSubItems(I_COL_SOUM_DISTRIB).Tag = lvwfournisseur.SelectedItem.Tag

            'Temps
785         itmNouveau.SubItems(I_COL_SOUM_TEMPS) = itmAncien.SubItems(I_COL_SOUM_TEMPS)

            'Si le temps n'est pas vide
790         If Trim$(itmNouveau.SubItems(I_COL_SOUM_TEMPS)) <> vbNullString Then
              'On calcul le temps * quantité pour la colonne montage
795           itmNouveau.SubItems(I_COL_SOUM_MONTAGE) = CDbl(Replace(itmNouveau.SubItems(I_COL_SOUM_TEMPS), ".", ",")) * CDbl(Replace(itmNouveau.Text, "*", vbNullString))
800         Else
805           itmNouveau.SubItems(I_COL_SOUM_MONTAGE) = vbNullString
810         End If

            'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
815         itmNouveau.SubItems(I_COL_SOUM_TOTAL) = Conversion(Round(CDbl(Replace(itmNouveau.Text, "*", "")) * CDbl(itmNouveau.SubItems(I_COL_SOUM_PRIX_NET)) * CDbl(m_sProfit), 2), MODE_ARGENT)

            'Pour le profit, c'est le prix total - (prix net * quantité)
820         itmNouveau.SubItems(I_COL_SOUM_PROFIT) = Conversion(Round(CDbl(itmNouveau.SubItems(I_COL_SOUM_TOTAL)) - (CDbl(Replace(itmNouveau.Text, "*", "") * CDbl(itmNouveau.SubItems(I_COL_SOUM_PRIX_NET)))), 2), MODE_ARGENT)

825         If Right$(txtNoProjSoum.Text, 2) >= "00" And Right$(txtNoProjSoum.Text, 2) <= "19" Then
              'Pour savoir lors de l'enregistrement qu'il faut le lier avec un extra
830           itmNouveau.ListSubItems(I_COL_SOUM_PROFIT).Tag = "RETOUR " & sExtra
835         End If

840         itmNouveau.SubItems(I_COL_SOUM_DATE_COMMANDE) = " "
845         itmNouveau.SubItems(I_COL_SOUM_DATE_REQUISE) = " "
850         itmNouveau.SubItems(I_COL_SOUM_NOM_COMMANDE) = " "
855         itmNouveau.SubItems(I_COL_SOUM_NO_SEQUENTIEL) = " "

860         If itmAncien.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_NOIR Then
865           itmNouveau.ListSubItems(I_COL_SOUM_DATE_COMMANDE).ForeColor = COLOR_NOIR
870           itmNouveau.ListSubItems(I_COL_SOUM_DATE_REQUISE).ForeColor = COLOR_NOIR
875           itmNouveau.ListSubItems(I_COL_SOUM_DESCR).ForeColor = COLOR_NOIR
880           itmNouveau.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = COLOR_NOIR
885           itmNouveau.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = COLOR_NOIR
890           itmNouveau.ListSubItems(I_COL_SOUM_ID).ForeColor = COLOR_NOIR
895           itmNouveau.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = COLOR_NOIR
900           itmNouveau.ListSubItems(I_COL_SOUM_MONTAGE).ForeColor = COLOR_NOIR
905           itmNouveau.ListSubItems(I_COL_SOUM_NO_SEQUENTIEL).ForeColor = COLOR_NOIR
910           itmNouveau.ListSubItems(I_COL_SOUM_NOM_COMMANDE).ForeColor = COLOR_NOIR
915           itmNouveau.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_NOIR
920           itmNouveau.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = COLOR_NOIR
925           itmNouveau.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = COLOR_NOIR
930           itmNouveau.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = COLOR_NOIR
935           itmNouveau.ListSubItems(I_COL_SOUM_TEMPS).ForeColor = COLOR_NOIR
940           itmNouveau.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = COLOR_NOIR
945         Else
950           itmNouveau.ListSubItems(I_COL_SOUM_DATE_COMMANDE).ForeColor = COLOR_BRUN
955           itmNouveau.ListSubItems(I_COL_SOUM_DATE_REQUISE).ForeColor = COLOR_BRUN
960           itmNouveau.ListSubItems(I_COL_SOUM_DESCR).ForeColor = COLOR_BRUN
965           itmNouveau.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = COLOR_BRUN
970           itmNouveau.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = COLOR_BRUN
975           itmNouveau.ListSubItems(I_COL_SOUM_ID).ForeColor = COLOR_BRUN
980           itmNouveau.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = COLOR_BRUN
985           itmNouveau.ListSubItems(I_COL_SOUM_MONTAGE).ForeColor = COLOR_BRUN
990           itmNouveau.ListSubItems(I_COL_SOUM_NO_SEQUENTIEL).ForeColor = COLOR_BRUN
995           itmNouveau.ListSubItems(I_COL_SOUM_NOM_COMMANDE).ForeColor = COLOR_BRUN
1000          itmNouveau.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_BRUN
1005          itmNouveau.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = COLOR_BRUN
1010          itmNouveau.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = COLOR_BRUN
1015          itmNouveau.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = COLOR_BRUN
1020          itmNouveau.ListSubItems(I_COL_SOUM_TEMPS).ForeColor = COLOR_BRUN
1025          itmNouveau.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = COLOR_BRUN
1030        End If

1035        Call CalculerTempsFabrication

            'Calcul des prix
1040        Call CalculerPrix
  
            'On cache le ListView
1045        frafournisseur.Visible = False

1050        m_bPieceInutile = False
  
            'Resélectionne le premier élément du listview
1055        If lvwSoumission.ListItems.count > 0 Then
1060          Call Deselect

1065          lvwSoumission.ListItems(1).Selected = True
1070        End If
1075      Else
1080        Call MsgBox("Quantité trop grande!", vbOKOnly, "Erreur")
1085      End If
1090    End If

1095    Exit Sub

AfficherErreur:

1100    woups "frmProjSoumElec", "ChoisirFournisseurMateriel", Err, Erl
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

40      woups "frmProjSoumElec", "lvwFournisseur_DblClick", Err, Erl
End Sub

Private Sub lvwPieces_DblClick()

5       On Error GoTo AfficherErreur

10      m_bPieceInutile = False
15      m_bRecherchePiece = False
20      m_bChangementFRS = False
        'On affiche lvwFournisseur selon la pièce choisie
        'Rempli les fournisseurs de la pièce choisie
25      Call AfficherListeFournisseurs
  
        'si le listview n'est pas vide
30      If lvwfournisseur.ListItems.count = 1 Then
35        If MsgBox("Il n'y a aucun fournisseur pour cette pièce!" & vbNewLine & "Voulez-vous en ajouter?", vbYesNo, "Erreur") = vbYes Then
40          Screen.MousePointer = vbHourglass
      
            'On ouvre le catalogue sur cet enregistrement
45          Call FrmCatalogueElec.AfficherForm(cmbPieces.Text, lvwPieces.SelectedItem.SubItems(I_COL_PIECES_MANUFACT), lvwPieces.SelectedItem.SubItems(I_COL_PIECES_NO_ITEM))
     
50          Screen.MousePointer = vbDefault
55        End If
60      End If

65      Exit Sub

AfficherErreur:

70      woups "frmProjSoumElec", "lvwPieces_DblClick", Err, Erl
End Sub

Private Sub AfficherListeFournisseurs()

5       On Error GoTo AfficherErreur

        'Méthode qui sert à afficher la liste des fournisseurs
        'Affiche le frame seulement s'il y a des items dans le ListView
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

55      woups "frmProjSoumElec", "AfficherListeFournisseurs", Err, Erl
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
70              If Trim$(m_sTexteRecherche) <> "" Then
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
250                 Else
255                   If KeyCode = vbKeyI Then
260                     If m_eType = TYPE_PROJET Then
265                       If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PROFIT).Tag <> "EXTRA" Then
270                         lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_ID) = InputBox("Quel est l'ID", , lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_ID))
275                       Else
280                         Call MsgBox("Cette commande doit être faite dans le projet " & lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PROVENANCE), vbOKOnly, "Erreur")
285                       End If
290                     End If
295                   End If
300                 End If
305               End If
310             End If
315           End If
320         End If
325       End If
330     End If

335     Exit Sub

AfficherErreur:

340     woups "frmProjSoumElec", "lvwSoumission_KeyDown", Err, Erl
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

95      woups "frmProjSoumElec", "FacturerDate", Err, Erl
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

95      woups "frmProjSoumElec", "FacturerNC", Err, Erl
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

85          Call Deselect

90          lvwSoumission.ListItems(iCompteur).Selected = True

95          Call lvwSoumission.SelectedItem.EnsureVisible

100         bTrouve = True

105         Exit For
110       End If
115     Next

120     If bTrouve = False Then
125       For iCompteur = 1 To iSelected - 1
130         If InStr(1, UCase(lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE)), UCase(sTexte)) > 0 Then
135           Call lvwSoumission.SetFocus

140           Call Deselect

145           lvwSoumission.ListItems(iCompteur).Selected = True

150           Call lvwSoumission.SelectedItem.EnsureVisible

155           bTrouve = True

160           Exit For
165         End If
170       Next
175     End If

180     If bTrouve = False Then
185       Call MsgBox("Aucun enregistrement trouvé!", vbOKOnly, "Erreur")
190     End If

195     Exit Sub

AfficherErreur:

200     woups "frmProjSoumElec", "RechercherPieceListViewSoumission", Err, Erl
End Sub

Private Sub EffacerItemListViewSoumission()

5       On Error GoTo AfficherErreur

10      Dim bSeulSS       As Boolean  'Pour savoir si c'est le seul enr. dans la sous-section
15      Dim bSeulS        As Boolean  'Pour savoir si c'est le seul enr. dans la section
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
                                        
                        'Il faut vérifier si c'est le seul enregistrement de la section. Si c'est le cas
                        'Il faut effacer la section en meme temps
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
          
                        'Si c'est le seul dans la sous-section, on efface la sous-section
430                     If bSeulSS = True Then
435                       Call lvwSoumission.ListItems.Remove(iIndex - 1)

440                       iCompteur = iCompteur - 1
445                     End If
  
                        'Si c'est le seul dans la section, on efface la section
450                     If bSeulS = True Then
455                       Call lvwSoumission.ListItems.Remove(iIndex - 2)

460                       iCompteur = iCompteur - 1
465                     End If
                   
                        'On recalcule le temps mécanique
470                     Call CalculerTempsFabrication

                        'On recalcule les prix
475                     Call CalculerPrix
480                   Else
485                     Call MsgBox("La pièce " & lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) & " doit être effacée dans le projet " & lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PROVENANCE), vbOKOnly, "Erreur")

490                     iCompteur = iCompteur + 1
495                   End If
500                 Else
505                   Call MsgBox("La pièce " & lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) & " ne peut pas être supprimée!", vbOKOnly, "Erreur")

510                   iCompteur = iCompteur + 1
515                 End If
520               Else
525                 iCompteur = iCompteur + 1
530               End If
535             Else
540               iCompteur = iCompteur + 1
545             End If
550           Else
555             iCompteur = iCompteur + 1
560           End If
565         Loop
570       Else
            'Cette ligne sert seulement à ne pas déselectionner et repositionner à la ligne 1 si l'utilisateur
            'décide de ne pas supprimer.
            'Le nom de la variable n'est pas significatif dans ce cas, mais c'est celle-ci qui est utilisé pour
            'désélectionner et remettre à la ligne 1
575         bPermission = False
580       End If
585     End If
        
        'Il faut resélectionner le premier à la fin
590     If lvwSoumission.ListItems.count > 0 Then
595       If bPermission = True Then
600         Call Deselect

605         lvwSoumission.ListItems(1).Selected = True
610       End If
615     End If

620     Exit Sub

AfficherErreur:

625     woups "frmProjSoumElec", "EffacerItemListViewSoumission", Err, Erl
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

45      woups "frmProjSoumElec", "AjouterSuppressionCollection", Err, Erl
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

65      For iCompteur = 1 To m_collNoItemSupp.count
70        Call rstBavard.AddNew

75        rstBavard.Fields("IDUser") = iNoEmploye
80        rstBavard.Fields("NoProjsoum") = txtNoProjSoum.Text
85        rstBavard.Fields("Type") = "E"
90        rstBavard.Fields("Qté") = m_collQteSupp(iCompteur)
95        rstBavard.Fields("No Item") = m_collNoItemSupp(iCompteur)
100       rstBavard.Fields("Date") = m_collDateSupp(iCompteur)
105       rstBavard.Fields("Heure") = m_collHeureSupp(iCompteur)

110       Call rstBavard.Update
115     Next

120     Call rstBavard.Close
125     Set rstBavard = Nothing

130     Exit Sub

AfficherErreur:

135     woups "frmProjSoumElec", "EnregistrerSuppression", Err, Erl
End Sub

Private Sub mvwDateRequise_GotFocus()
        
5       On Error GoTo AfficherErreur

10      m_bMonthViewHasFocus = True

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElec", "mvwDateRequise_GotFocus", Err, Erl
End Sub

Private Sub tmrTemps_Timer()
  
5       On Error GoTo AfficherErreur

10      If lblPasTemps.Visible = True Then
15        lblPasTemps.Visible = False
20      Else
25        lblPasTemps.Visible = True
30      End If

35      Exit Sub

AfficherErreur:

40      woups "frmProjSoumElec", "tmrTemps_Timer", Err, Erl
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

40      woups "frmProjSoumElec", "txtCheminPhotos_KeyDown", Err, Erl
End Sub

Private Sub txtPrixManuel_Change()

5       On Error GoTo AfficherErreur
        
        'Si le texte change, il faut recalculer les prix
10      Call CalculerPrix

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElec", "txtManuel_Change", Err, Erl
End Sub

Private Sub cmdAnnulerPrix_Click()

5       On Error GoTo AfficherErreur

10      fraPrixPiece.Visible = False

15      m_bMauvaisPrix = False

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumElec", "cmdAnnulerPrix_Click", Err, Erl
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
40      Dim sQuantite    As String
45      Dim sPiece       As String
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
170       If Trim$(txtPrixNet.Text) <> vbNullString Then
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

245       Set itmAvant = lvwSoumission.ListItems(CInt(fraPrixPiece.Tag))
250       Set itmSoum = lvwSoumission.ListItems.Add(CInt(fraPrixPiece.Tag) + 1)

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

          'Temps
340       itmSoum.SubItems(I_COL_SOUM_TEMPS) = itmAvant.SubItems(I_COL_SOUM_TEMPS)

          'Si le temps n'est pas vide
345       If Trim$(itmSoum.SubItems(I_COL_SOUM_TEMPS)) <> vbNullString Then
            'On calcul le temps * quantité pour la colonne montage
350         itmSoum.SubItems(I_COL_SOUM_MONTAGE) = CDbl(Replace(itmSoum.SubItems(I_COL_SOUM_TEMPS), ".", ",")) * CDbl(Replace(itmSoum.Text, "*", vbNullString))
355       Else
360         itmSoum.SubItems(I_COL_SOUM_MONTAGE) = vbNullString
365       End If

          'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
370       itmSoum.SubItems(I_COL_SOUM_TOTAL) = "-" & itmAvant.SubItems(I_COL_SOUM_TOTAL)

          'Pour le profit, c'est le prix total - (prix net * quantité)
375       itmSoum.SubItems(I_COL_SOUM_PROFIT) = "-" & itmAvant.SubItems(I_COL_SOUM_PROFIT)

          'Ajout de l'enregistrement avec le nouveau prix
380       Set itmSoum = lvwSoumission.ListItems.Add(CInt(fraPrixPiece.Tag) + 2)

385       itmSoum.Checked = itmAvant.Checked

          'Quantité
390       itmSoum.Text = sQuantite

          'On met l'id de la section dans le tag du listItem
395       itmSoum.Tag = itmAvant.Tag

          'No d'item
400       itmSoum.SubItems(I_COL_SOUM_PIECE) = itmAvant.SubItems(I_COL_SOUM_PIECE)

405       itmSoum.ListSubItems(I_COL_SOUM_PIECE).ForeColor = lColor

          'On met le nom de la sous-section dans le tag du no d'item
410       itmSoum.ListSubItems(I_COL_SOUM_PIECE).Tag = itmAvant.ListSubItems(I_COL_SOUM_PIECE).Tag

          'On met la description en francais dans la colonne et la description en anglais
          'dans le tag
415       itmSoum.SubItems(I_COL_SOUM_DESCR) = itmAvant.SubItems(I_COL_SOUM_DESCR)
420       itmSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = itmAvant.ListSubItems(I_COL_SOUM_DESCR).Tag

425       itmSoum.ListSubItems(I_COL_SOUM_DESCR).ForeColor = lColor

          'On met le nom du fabricant dans la col-nne et l'ordre de la section dans le tag
430       itmSoum.SubItems(I_COL_SOUM_MANUFACT) = itmAvant.SubItems(I_COL_SOUM_MANUFACT)
435       itmSoum.ListSubItems(I_COL_SOUM_MANUFACT).Tag = itmAvant.ListSubItems(I_COL_SOUM_MANUFACT).Tag

440       itmSoum.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = lColor

445       If bPrixSpecial = False Then
450         If optUSA.Value = True Then
455           itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixList.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
460         Else
465           If optSpain.Value = True Then
470             itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixList.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
475           Else
480             itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(txtPrixList.Text, MODE_ARGENT, 4)
485           End If
490         End If

495         itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = txtPrixList.Text
       
500         itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = lColor
       
            'Escompte
505         If mskEscompte.Text <> vbNullString Then
510           itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion(mskEscompte.Text, MODE_POURCENT)
515         Else
520           itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion("0", MODE_POURCENT)
525         End If

530         itmSoum.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = lColor

            'Prix net
535         If optUSA.Value = True Then
540           itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixNet.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
545         Else
550           If optSpain.Value = True Then
555             itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixNet.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
560           Else
565             itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(txtPrixNet.Text, MODE_ARGENT, 4)
570           End If
575         End If

580         itmSoum.ListSubItems(I_COL_SOUM_PRIX_NET).Tag = itmAvant.ListSubItems(I_COL_SOUM_PRIX_NET).Tag

585         itmSoum.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = lColor
590       Else
595         If optUSA.Value = True Then
600           itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
605         Else
610           If optSpain.Value = True Then
615             itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
620           Else
625             itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(txtPrixSpecial.Text, MODE_ARGENT, 4)
630           End If
635         End If

640         itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = txtPrixSpecial.Text

645         itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = lColor

650         itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion("0", MODE_POURCENT)

655         itmSoum.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = lColor

660         If optUSA.Value = True Then
665           itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
670         Else
675           If optSpain.Value = True Then
680             itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
685           Else
690             itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(txtPrixSpecial.Text, MODE_ARGENT, 4)
695           End If
700         End If

705         itmSoum.ListSubItems(I_COL_SOUM_PRIX_NET).Tag = itmAvant.ListSubItems(I_COL_SOUM_PRIX_NET).Tag

710         itmSoum.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = lColor
715       End If

          'On met le fournisseur dans la colonne et l'id dans le tag
720       itmSoum.SubItems(I_COL_SOUM_DISTRIB) = itmAvant.SubItems(I_COL_SOUM_DISTRIB)
725       itmSoum.ListSubItems(I_COL_SOUM_DISTRIB).Tag = itmAvant.ListSubItems(I_COL_SOUM_DISTRIB).Tag

730       itmSoum.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = lColor
    
          'Temps
735       itmSoum.SubItems(I_COL_SOUM_TEMPS) = itmAvant.SubItems(I_COL_SOUM_TEMPS)

740       itmSoum.ListSubItems(I_COL_SOUM_TEMPS).ForeColor = lColor
    
          'Si le temps n'est pas vide
745       If Trim$(itmSoum.SubItems(I_COL_SOUM_TEMPS)) <> vbNullString Then
            'On calcul le temps * quantité pour la colonne montage
750         itmSoum.SubItems(I_COL_SOUM_MONTAGE) = CDbl(Replace(itmSoum.SubItems(I_COL_SOUM_TEMPS), ".", ",")) * CDbl(Replace(itmSoum.Text, "*", vbNullString))
755       Else
760         itmSoum.SubItems(I_COL_SOUM_MONTAGE) = vbNullString
765       End If
      
770       itmSoum.ListSubItems(I_COL_SOUM_MONTAGE).ForeColor = lColor

          'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
775       itmSoum.SubItems(I_COL_SOUM_TOTAL) = Conversion(CStr(Round(Replace(itmSoum.Text, "*", vbNullString) * itmSoum.SubItems(I_COL_SOUM_PRIX_NET) * CSng(m_sProfit), 2)), MODE_ARGENT)
      
780       itmSoum.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = lColor

785       If optUSA.Value = True Then
790         itmSoum.ListSubItems(I_COL_SOUM_TOTAL).Tag = "USA"
795       Else
800         If optSpain.Value = True Then
805           itmSoum.ListSubItems(I_COL_SOUM_TOTAL).Tag = "SPA"
810         Else
815           itmSoum.ListSubItems(I_COL_SOUM_TOTAL).Tag = "CAN"
820         End If
825       End If
     
          'Pour le profit, c'est le prix total - (prix net * quantité)
830       itmSoum.SubItems(I_COL_SOUM_PROFIT) = Conversion(CStr(Round(itmSoum.SubItems(I_COL_SOUM_TOTAL) - (itmSoum.SubItems(I_COL_SOUM_PRIX_NET) * Replace(itmSoum.Text, "*", vbNullString)), 2)), MODE_ARGENT)

835       itmSoum.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = lColor

840       itmSoum.SubItems(I_COL_SOUM_COMMENTAIRE) = itmAvant.SubItems(I_COL_SOUM_COMMENTAIRE)
845       itmSoum.ListSubItems(I_COL_SOUM_COMMENTAIRE).ForeColor = lColor

850       If m_eType = TYPE_PROJET Then
855         itmSoum.SubItems(I_COL_SOUM_DATE_COMMANDE) = itmAvant.SubItems(I_COL_SOUM_DATE_COMMANDE)
860         itmSoum.ListSubItems(I_COL_SOUM_DATE_COMMANDE).ForeColor = lColor

865         itmSoum.SubItems(I_COL_SOUM_DATE_REQUISE) = itmAvant.SubItems(I_COL_SOUM_DATE_REQUISE)
870         itmSoum.ListSubItems(I_COL_SOUM_DATE_REQUISE).ForeColor = lColor

875         itmSoum.SubItems(I_COL_SOUM_NOM_COMMANDE) = itmAvant.SubItems(I_COL_SOUM_NOM_COMMANDE)
880         itmSoum.ListSubItems(I_COL_SOUM_NOM_COMMANDE).ForeColor = lColor

885         itmSoum.SubItems(I_COL_SOUM_NO_SEQUENTIEL) = itmAvant.SubItems(I_COL_SOUM_NO_SEQUENTIEL)
890         itmSoum.ListSubItems(I_COL_SOUM_NO_SEQUENTIEL).ForeColor = lColor

895         itmSoum.SubItems(I_COL_SOUM_ID) = itmAvant.SubItems(I_COL_SOUM_ID)
900         itmSoum.ListSubItems(I_COL_SOUM_ID).ForeColor = lColor
         
905         itmSoum.SubItems(I_COL_SOUM_FACTURATION) = itmAvant.SubItems(I_COL_SOUM_FACTURATION)

910         If itmSoum.SubItems(I_COL_SOUM_FACTURATION) <> "" Then
915           itmSoum.ListSubItems(I_COL_SOUM_FACTURATION).Tag = itmAvant.ListSubItems(I_COL_SOUM_FACTURATION)
920         End If

925         itmSoum.ListSubItems(I_COL_SOUM_FACTURATION).ForeColor = lColor
930       End If

935       If itmAvant.SubItems(I_COL_SOUM_COMMENTAIRE) <> "" Then
940         itmSoum.ListSubItems(I_COL_SOUM_COMMENTAIRE).ForeColor = itmAvant.ListSubItems(I_COL_SOUM_COMMENTAIRE).ForeColor

945         itmAvant.ListSubItems(I_COL_SOUM_COMMENTAIRE).ForeColor = vbBlack
950       End If

955       If m_eType = TYPE_PROJET Then
960         If itmAvant.SubItems(I_COL_SOUM_DATE_COMMANDE) <> "" Then
965           itmAvant.ListSubItems(I_COL_SOUM_DATE_COMMANDE).ForeColor = vbBlack
970         End If

975         If itmAvant.SubItems(I_COL_SOUM_DATE_REQUISE) <> "" Then
980           itmAvant.ListSubItems(I_COL_SOUM_DATE_REQUISE).ForeColor = vbBlack
985         End If

990         If itmAvant.SubItems(I_COL_SOUM_NOM_COMMANDE) <> "" Then
995           itmAvant.ListSubItems(I_COL_SOUM_NOM_COMMANDE).ForeColor = vbBlack
1000        End If

1005        If itmAvant.SubItems(I_COL_SOUM_NO_SEQUENTIEL) <> "" Then
1010          itmAvant.ListSubItems(I_COL_SOUM_NO_SEQUENTIEL).ForeColor = vbBlack
1015        End If

1020        If itmAvant.SubItems(I_COL_SOUM_FACTURATION) <> "" Then
1025          itmAvant.ListSubItems(I_COL_SOUM_FACTURATION).ForeColor = vbBlack
1030        End If

1035        If itmAvant.SubItems(I_COL_SOUM_ID) <> "" Then
1040          itmAvant.ListSubItems(I_COL_SOUM_ID).ForeColor = vbBlack
1045        End If

1050        If itmAvant.SubItems(I_COL_SOUM_PROVENANCE) <> "" Then
1055          itmAvant.ListSubItems(I_COL_SOUM_PROVENANCE).ForeColor = vbBlack
1060        End If
1065      End If

1070      itmSoum.ListSubItems(I_COL_SOUM_DESCR).ForeColor = itmAvant.ListSubItems(I_COL_SOUM_DESCR).ForeColor
1075      itmSoum.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = itmAvant.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor
1080      itmSoum.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = itmAvant.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor
1085      itmSoum.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = itmAvant.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor
1090      itmSoum.ListSubItems(I_COL_SOUM_MONTAGE).ForeColor = itmAvant.ListSubItems(I_COL_SOUM_MONTAGE).ForeColor
1095      itmSoum.ListSubItems(I_COL_SOUM_PIECE).ForeColor = itmAvant.ListSubItems(I_COL_SOUM_PIECE).ForeColor
1100      itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = itmAvant.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor
1105      itmSoum.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = itmAvant.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor
1110      itmSoum.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = itmAvant.ListSubItems(I_COL_SOUM_PROFIT).ForeColor
1115      itmSoum.ListSubItems(I_COL_SOUM_TEMPS).ForeColor = itmAvant.ListSubItems(I_COL_SOUM_TEMPS).ForeColor
1120      itmSoum.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = itmAvant.ListSubItems(I_COL_SOUM_TOTAL).ForeColor

1125      itmAvant.ListSubItems(I_COL_SOUM_DESCR).ForeColor = vbBlack
1130      itmAvant.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = vbBlack
1135      itmAvant.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = vbBlack
1140      itmAvant.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = vbBlack
1145      itmAvant.ListSubItems(I_COL_SOUM_MONTAGE).ForeColor = vbBlack
1150      itmAvant.ListSubItems(I_COL_SOUM_PIECE).ForeColor = vbBlack
1155      itmAvant.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = vbBlack
1160      itmAvant.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = vbBlack
1165      itmAvant.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = vbBlack
1170      itmAvant.ListSubItems(I_COL_SOUM_TEMPS).ForeColor = vbBlack
1175      itmAvant.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = vbBlack

1180      Call CalculerTempsFabrication

          'Resélectionne le premier élément du listview
1185      If lvwSoumission.ListItems.count > 0 Then
1190        Call Deselect

1195        lvwSoumission.ListItems(1).Selected = True
1200      End If
          
1205      m_bMauvaisPrix = False

1210      cmbfrs.Locked = False

1215      Call lvwSoumission.Refresh
1220    Else
1225      sPiece = lvwSoumission.ListItems(CInt(fraPrixPiece.Tag)).SubItems(I_COL_SOUM_PIECE)

1230      For iCompteur = 1 To lvwSoumission.ListItems.count
1235        If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) = sPiece And lvwSoumission.ListItems(iCompteur).ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_MAGENTA Then
1240          Set itmSoum = lvwSoumission.ListItems(iCompteur)

1245          itmSoum.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_NOIR
1250          itmSoum.ListSubItems(I_COL_SOUM_DESCR).ForeColor = COLOR_NOIR
1255          itmSoum.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = COLOR_NOIR
1260          itmSoum.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = COLOR_NOIR
1265          itmSoum.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = COLOR_NOIR
1270          itmSoum.ListSubItems(I_COL_SOUM_MONTAGE).ForeColor = COLOR_NOIR
1275          itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = COLOR_NOIR
1280          itmSoum.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = COLOR_NOIR
1285          itmSoum.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = COLOR_NOIR
1290          itmSoum.ListSubItems(I_COL_SOUM_TEMPS).ForeColor = COLOR_NOIR
1295          itmSoum.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = COLOR_NOIR

1300          If itmSoum.SubItems(I_COL_SOUM_COMMENTAIRE) <> "" Then
1305            itmSoum.ListSubItems(I_COL_SOUM_COMMENTAIRE).ForeColor = COLOR_NOIR
1310          End If

1315          Call lvwSoumission.Refresh
  
1320          If bPrixSpecial = False Then
                'Prix listé
1325            If optUSA.Value = True Then
1330              itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixList.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
1335            Else
1340              If optSpain.Value = True Then
1345                itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixList.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
1350              Else
1355                itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(txtPrixList.Text, MODE_ARGENT, 4)
1360              End If
1365            End If

1370            itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = txtPrixList.Text
        
                'Escompte
1375            If mskEscompte.Text <> vbNullString Then
1380              itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion(mskEscompte.Text, MODE_POURCENT)
1385            Else
1390              itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion("0", MODE_POURCENT)
1395            End If

                'Prix net
1400            If optUSA.Value = True Then
1405              itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixNet.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
1410            Else
1415              If optSpain.Value = True Then
1420                itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixNet.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
1425              Else
1430                itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(txtPrixNet.Text, MODE_ARGENT, 4)
1435              End If
1440            End If
1445          Else
1450            If optUSA.Value = True Then
1455              itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
1460            Else
1465              If optSpain.Value = True Then
1470                itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
1475              Else
1480                itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(txtPrixSpecial.Text, MODE_ARGENT, 4)
1485              End If
1490            End If

1495            itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = txtPrixSpecial.Text
         
1500            itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion("0", MODE_POURCENT)

1505            If optUSA.Value = True Then
1510              itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
1515            Else
1520              If optSpain.Value = True Then
1525                itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
1530              Else
1535                itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(txtPrixSpecial.Text, MODE_ARGENT, 4)
1540              End If
1545            End If
1550          End If

              'On met le fournisseur dans la colonne et l'id dans le tag
1555          itmSoum.SubItems(I_COL_SOUM_DISTRIB) = cmbfrs.LIST(cmbfrs.ListIndex)
    
1560          itmSoum.ListSubItems(I_COL_SOUM_DISTRIB).Tag = cmbfrs.ItemData(cmbfrs.ListIndex)
           
              'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
1565          itmSoum.SubItems(I_COL_SOUM_TOTAL) = Conversion(CStr(Round(Replace(itmSoum.Text, "*", "") * itmSoum.SubItems(I_COL_SOUM_PRIX_NET) * CSng(m_sProfit), 2)), MODE_ARGENT)

1570          If optUSA.Value = True Then
1575            itmSoum.ListSubItems(I_COL_SOUM_TOTAL).Tag = "USA"
1580          Else
1585            If optSpain.Value = True Then
1590              itmSoum.ListSubItems(I_COL_SOUM_TOTAL).Tag = "SPA"
1595            Else
1600              itmSoum.ListSubItems(I_COL_SOUM_TOTAL).Tag = "CAN"
1605            End If
1610          End If
      
              'Pour le profit, c'est le prix total - (prix net * quantité)
1615          itmSoum.SubItems(I_COL_SOUM_PROFIT) = Conversion(CStr(Round(itmSoum.SubItems(I_COL_SOUM_TOTAL) - (itmSoum.SubItems(I_COL_SOUM_PRIX_NET) * Replace(itmSoum.Text, "*", "")), 2)), MODE_ARGENT)
1620        End If
1625      Next
1630    End If

1635    Call ModifierPrixCatalogue

1640    fraPrixPiece.Visible = False

1645    Call CalculerPrix

1650    Exit Sub

AfficherErreur:

1655    woups "frmProjSoumElec", "cmdOKPrix_Click", Err, Erl
End Sub

Private Sub RemplirComboFournisseur()

5       On Error GoTo AfficherErreur

10      Dim rstFRS    As ADODB.Recordset
15      Dim iCompteur As Integer
20      Dim bExiste   As Boolean

25      Set rstFRS = New ADODB.Recordset

        'Il faut vider le combo avant de le remplir
30      Call cmbfrs.Clear

35      Call rstFRS.Open("SELECT GRB_PiecesFRS.*, GRB_Fournisseur.NomFournisseur FROM GRB_PiecesFRS INNER JOIN GRB_Fournisseur ON GRB_PiecesFRS.IDFRS = GRB_Fournisseur.IDFRS WHERE PIECE = '" & Replace(lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE), "'", "''") & "' ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)

        'Tant que ce n'est pas la fin des enregistrements
40      Do While Not rstFRS.EOF
45        bExiste = False

50        For iCompteur = 0 To cmbfrs.ListCount - 1
55          If cmbfrs.ItemData(iCompteur) = rstFRS.Fields("IDFRS") Then
60            bExiste = True

65            Exit For
70          End If
75        Next

80        If bExiste = False Then
85          Call cmbfrs.AddItem(rstFRS.Fields("NomFournisseur"))

90          cmbfrs.ItemData(cmbfrs.newIndex) = rstFRS.Fields("IDFRS")
95        End If

100       Call rstFRS.MoveNext
105     Loop

110     Call rstFRS.Close
115     Set rstFRS = Nothing

120     Exit Sub

AfficherErreur:

125     woups "frmProjSoumElec", "RemplirComboFournisseur", Err, Erl
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

60      woups "frmProjSoumElec", "txtPrixList_LostFocus", Err, Erl
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

45      woups "frmProjSoumElec", "txtPrixNet_Change", Err, Erl

End Sub

Private Sub txtPrixNet_GotFocus()

5       On Error GoTo AfficherErreur

        'Si le prix net prend le focus
10      Call CalculerPrixNet

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElec", "txtPrixNet_GotFocus", Err, Erl
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

105     woups "frmProjSoumElec", "CalculerPrixNet", Err, Erl
End Sub

Private Sub txtPrixNet_LostFocus()

5       On Error GoTo AfficherErreur

10      txtPrixNet.Text = Replace(txtPrixNet.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElec", "txtPrixNet_LostFocus", Err, Erl
End Sub

Private Sub ViderChamps_frs()

5      On Error GoTo AfficherErreur

        'Vide les champs pieces
10      txtPrixList.Text = vbNullString
15      mskEscompte.Text = vbNullString
20      txtPrixNet.Text = vbNullString
25      txtPrixSpecial.Text = vbNullString
  
30      optCAN.Value = True

35      Call AfficherDrapeau

40      Exit Sub

AfficherErreur:

45      woups "frmProjSoumElec", "ViderChamps_frs", Err, Erl
End Sub

Private Sub ModifierPrixCatalogue()
        'Enregistrement du prix de la pièce
        
5       On Error GoTo AfficherErreur

10      Dim rstPrix     As ADODB.Recordset
15      Dim dblPrixList As Double
20      Dim dblEscompte As Double
25      Dim dblPrixNet  As Double
                                
30      If Trim$(txtPrixList.Text) <> "" Then
35        dblPrixList = CDbl(txtPrixList.Text)
40      Else
45        dblPrixList = 0
50      End If
        
55      If mskEscompte.Text <> vbNullString Then
60        dblEscompte = CDbl(mskEscompte.Text)
65      Else
70        dblEscompte = 0
75      End If
        
80      If Trim$(txtPrixNet.Text) <> "" Then
85        dblPrixNet = CDbl(txtPrixNet.Text)
90      Else
95        dblPrixNet = CDbl(txtPrixSpecial.Text)
100     End If
                                                
105     Set rstPrix = New ADODB.Recordset
                                                
110     If txtPrixNet.Enabled = True Then
          'Ouverture du recordset
115       Call rstPrix.Open("SELECT * FROM GRB_PiecesFRS WHERE PIECE = '" & Replace(lvwSoumission.ListItems(CInt(fraPrixPiece.Tag)).SubItems(I_COL_SOUM_PIECE), "'", "''") & "' AND IDFRS = " & cmbfrs.ItemData(cmbfrs.ListIndex) & " AND PRIX_NET <> ''", g_connData, adOpenDynamic, adLockOptimistic)

120       If rstPrix.EOF Then
125         Call rstPrix.AddNew

130         rstPrix.Fields("PIECE").Value = lvwSoumission.ListItems(CInt(fraPrixPiece.Tag)).SubItems(I_COL_SOUM_PIECE)
135         rstPrix.Fields("IDFRS").Value = cmbfrs.ItemData(cmbfrs.ListIndex)
140       End If

145       rstPrix.Fields("PRIX_LIST") = dblPrixList
150       rstPrix.Fields("ESCOMPTE") = dblEscompte
155       rstPrix.Fields("PRIX_NET") = dblPrixNet
160       rstPrix.Fields("PRIX_SP") = ""
165     Else
170       Call rstPrix.Open("SELECT * FROM GRB_PiecesFRS WHERE PIECE = '" & Replace(lvwSoumission.ListItems(CInt(fraPrixPiece.Tag)).SubItems(I_COL_SOUM_PIECE), "'", "''") & "' AND IDFRS = " & cmbfrs.ItemData(cmbfrs.ListIndex) & " AND PRIX_SP <> ''", g_connData, adOpenDynamic, adLockOptimistic)

175       If rstPrix.EOF Then
180         Call rstPrix.AddNew

185         rstPrix.Fields("PIECE").Value = lvwSoumission.ListItems(CInt(fraPrixPiece.Tag)).SubItems(I_COL_SOUM_PIECE)
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
  
270     rstPrix.Fields("Type") = "E"

275     rstPrix.Fields("ENTRER_PAR") = g_sInitiale

280     rstPrix.Fields("Date") = ConvertDate(Date)

285     Call rstPrix.Update
  
290     Call rstPrix.Close
295     Set rstPrix = Nothing
       
300     Exit Sub

AfficherErreur:

305     woups "frmProjSoumElec", "ModifierPrixCatalogue", Err, Erl
End Sub

Private Sub optCAN_Click()

5       On Error GoTo AfficherErreur

        'dependant la devise, affiche le drapeau
10      Call AfficherDrapeau

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElec", "optCAN_Click", Err, Erl
End Sub
            
Private Sub AfficherDrapeau()

5       On Error GoTo AfficherErreur

        '''''''''''''''''''''''''''''
        'dependant la devise, affiche le drapeau
        '''''''''''''''''''''''''''''''''''''
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

90      woups "frmProjSoumElec", "AfficherDrapeau", Err, Erl
End Sub

Private Sub optSpain_Click()

5       On Error GoTo AfficherErreur

        'dependant la devise, affiche le drapeau
10      Call AfficherDrapeau

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElec", "optSpain_Click", Err, Erl
End Sub

Private Sub optUSA_Click()

5       On Error GoTo AfficherErreur

        'Dépendant la devise, affiche le drapeau
10      Call AfficherDrapeau

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElec", "optUSA_Click", Err, Erl
End Sub

Private Sub mskEscompte_GotFocus()

5       On Error GoTo AfficherErreur

        'Quand le maskEdit prend le focus, on set le masque
10      If mskEscompte.Enabled = True Then
15        mskEscompte.mask = "0,####"
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmProjSoumElec", "mskEscompte_GotFocus", Err, Erl
End Sub

Private Sub mskEscompte_LostFocus()

5       On Error GoTo AfficherErreur

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

40      woups "frmProjSoumElec", "mskEscompte_LostFocus", Err, Erl
End Sub

Private Function VerifierSiOuvert(ByRef sUser As String) As Boolean
        'Vérifie si le projet ou la soumission n'est pas en modification
        'par un autre utilisateur sur un autre ordinateur
5       On Error GoTo AfficherErreur

10      Dim rstProjSoum   As ADODB.Recordset
15      Dim bModification As Boolean

20      Set rstProjSoum = New ADODB.Recordset

25      If m_eType = TYPE_PROJET Then
30        Call rstProjSoum.Open("SELECT Modification, Par FROM GRB_ProjetElec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
35      Else
40        Call rstProjSoum.Open("SELECT Modification, Par FROM GRB_SoumissionElec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
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

105     woups "frmProjSoumElec", "VerifierSiOuvert", Err, Erl
End Function

Private Sub OuvrirProjSoum(ByVal bOuvrir As Boolean)
        'Remplis ou vide les champs Modification et Par
5       On Error GoTo AfficherErreur

10      Dim rstProjSoum As ADODB.Recordset

15      Set rstProjSoum = New ADODB.Recordset

20      rstProjSoum.CursorLocation = adUseServer

25      If m_eType = TYPE_PROJET Then
30        Call rstProjSoum.Open("SELECT Modification, Par FROM GRB_ProjetElec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
35      Else
40        Call rstProjSoum.Open("SELECT Modification, Par FROM GRB_SoumissionElec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
45      End If

50      Do While Not rstProjSoum.EOF
55        If bOuvrir = True Then
60          rstProjSoum.Fields("Modification") = True
65          rstProjSoum.Fields("Par") = g_sEmploye
70        Else
75          rstProjSoum.Fields("Modification") = False
80          rstProjSoum.Fields("Par") = ""
85        End If

90        Call rstProjSoum.Update
          
95        Call rstProjSoum.MoveNext
100     Loop

105     Call rstProjSoum.Close
110     Set rstProjSoum = Nothing

115     Exit Sub

AfficherErreur:

120     woups "frmProjSoumElec", "OuvrirProjSoum", Err, Erl
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

105       Call rstProjet.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sExtra & "'", g_connData, adOpenDynamic, adLockOptimistic)

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

        'On met le nom du fabricant dans la col-nne et l'ordre de la section dans le tag
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

        'Temps
245     itmAnnulation.SubItems(I_COL_SOUM_TEMPS) = itmAvant.SubItems(I_COL_SOUM_TEMPS)

        'Si le temps n'est pas vide
250     If Trim$(itmAnnulation.SubItems(I_COL_SOUM_TEMPS)) <> vbNullString Then
          'On calcul le temps * quantité pour la colonne montage
255       itmAnnulation.SubItems(I_COL_SOUM_MONTAGE) = CDbl(Replace(itmAnnulation.SubItems(I_COL_SOUM_TEMPS), ".", ",")) * CDbl(Replace(itmAnnulation.Text, "*", vbNullString))
260     Else
265       itmAnnulation.SubItems(I_COL_SOUM_MONTAGE) = vbNullString
270     End If

        'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
275     itmAnnulation.SubItems(I_COL_SOUM_TOTAL) = "-" & itmAvant.SubItems(I_COL_SOUM_TOTAL)

        'Pour le profit, c'est le prix total - (prix net * quantité)
280     itmAnnulation.SubItems(I_COL_SOUM_PROFIT) = "-" & itmAvant.SubItems(I_COL_SOUM_PROFIT)

285     If Right$(txtNoProjSoum.Text, 2) >= "00" And Right$(txtNoProjSoum.Text, 2) <= "19" Then
          'Pour savoir lors de l'enregistremenet qu'il faut le lier avec un extra
290       itmAnnulation.ListSubItems(I_COL_SOUM_PROFIT).Tag = "ANNULATION " & sExtra
295     End If

300     itmAnnulation.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_VERT_FORET
305     itmAnnulation.ListSubItems(I_COL_SOUM_DESCR).ForeColor = COLOR_VERT_FORET
310     itmAnnulation.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = COLOR_VERT_FORET
315     itmAnnulation.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = COLOR_VERT_FORET
320     itmAnnulation.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = COLOR_VERT_FORET
325     itmAnnulation.ListSubItems(I_COL_SOUM_MONTAGE).ForeColor = COLOR_VERT_FORET
330     itmAnnulation.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = COLOR_VERT_FORET
335     itmAnnulation.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = COLOR_VERT_FORET
340     itmAnnulation.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = COLOR_VERT_FORET
345     itmAnnulation.ListSubItems(I_COL_SOUM_TEMPS).ForeColor = COLOR_VERT_FORET
350     itmAnnulation.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = COLOR_VERT_FORET
                      
355     itmAnnulation.ListSubItems(I_COL_SOUM_PIECE).Bold = True
360     itmAnnulation.ListSubItems(I_COL_SOUM_DESCR).Bold = True
365     itmAnnulation.ListSubItems(I_COL_SOUM_DISTRIB).Bold = True
370     itmAnnulation.ListSubItems(I_COL_SOUM_ESCOMPTE).Bold = True
375     itmAnnulation.ListSubItems(I_COL_SOUM_MANUFACT).Bold = True
380     itmAnnulation.ListSubItems(I_COL_SOUM_MONTAGE).Bold = True
385     itmAnnulation.ListSubItems(I_COL_SOUM_PRIX_LIST).Bold = True
390     itmAnnulation.ListSubItems(I_COL_SOUM_PRIX_NET).Bold = True
395     itmAnnulation.ListSubItems(I_COL_SOUM_PROFIT).Bold = True
400     itmAnnulation.ListSubItems(I_COL_SOUM_TEMPS).Bold = True
405     itmAnnulation.ListSubItems(I_COL_SOUM_TOTAL).Bold = True

410     itmAvant.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_VERT_FORET
415     itmAvant.ListSubItems(I_COL_SOUM_DESCR).ForeColor = COLOR_VERT_FORET
420     itmAvant.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = COLOR_VERT_FORET
425     itmAvant.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = COLOR_VERT_FORET
430     itmAvant.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = COLOR_VERT_FORET
435     itmAvant.ListSubItems(I_COL_SOUM_MONTAGE).ForeColor = COLOR_VERT_FORET
440     itmAvant.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = COLOR_VERT_FORET
445     itmAvant.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = COLOR_VERT_FORET
450     itmAvant.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = COLOR_VERT_FORET
455     itmAvant.ListSubItems(I_COL_SOUM_TEMPS).ForeColor = COLOR_VERT_FORET
460     itmAvant.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = COLOR_VERT_FORET
465     itmAvant.ListSubItems(I_COL_SOUM_DATE_COMMANDE).ForeColor = COLOR_VERT_FORET
470     itmAvant.ListSubItems(I_COL_SOUM_DATE_REQUISE).ForeColor = COLOR_VERT_FORET
475     itmAvant.ListSubItems(I_COL_SOUM_NOM_COMMANDE).ForeColor = COLOR_VERT_FORET
480     itmAvant.ListSubItems(I_COL_SOUM_NO_SEQUENTIEL).ForeColor = COLOR_VERT_FORET

485     itmAvant.ListSubItems(I_COL_SOUM_PIECE).Bold = True
490     itmAvant.ListSubItems(I_COL_SOUM_DESCR).Bold = True
495     itmAvant.ListSubItems(I_COL_SOUM_DISTRIB).Bold = True
500     itmAvant.ListSubItems(I_COL_SOUM_ESCOMPTE).Bold = True
505     itmAvant.ListSubItems(I_COL_SOUM_MANUFACT).Bold = True
510     itmAvant.ListSubItems(I_COL_SOUM_MONTAGE).Bold = True
515     itmAvant.ListSubItems(I_COL_SOUM_PRIX_LIST).Bold = True
520     itmAvant.ListSubItems(I_COL_SOUM_PRIX_NET).Bold = True
525     itmAvant.ListSubItems(I_COL_SOUM_PROFIT).Bold = True
530     itmAvant.ListSubItems(I_COL_SOUM_TEMPS).Bold = True
535     itmAvant.ListSubItems(I_COL_SOUM_TOTAL).Bold = True
540     itmAvant.ListSubItems(I_COL_SOUM_DATE_COMMANDE).Bold = True
545     itmAvant.ListSubItems(I_COL_SOUM_DATE_REQUISE).Bold = True
550     itmAvant.ListSubItems(I_COL_SOUM_NOM_COMMANDE).Bold = True
555     itmAvant.ListSubItems(I_COL_SOUM_NO_SEQUENTIEL).Bold = True

560     Call lvwSoumission.Refresh

565     Call CalculerPrix

570     Exit Sub

AfficherErreur:

575     woups "frmProjSoumElec", "AnnulerCommande", Err, Erl
End Sub

Private Sub cmdEffacerForfait_Click()

5       On Error GoTo AfficherErreur

10      txtForfait.Text = ""
15      lblForfaitInitiale.Caption = ""

20      Exit Sub

AfficherErreur:

25      woups "frmProjSoumElec", "cmdEffacerForfait_Click", Err, Erl
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

180         m_arr_tyCopie(iIndex).sTemps = itmCopier.SubItems(I_COL_SOUM_TEMPS)

185         m_arr_tyCopie(iIndex).sMontage = itmCopier.SubItems(I_COL_SOUM_MONTAGE)

190         m_arr_tyCopie(iIndex).sTotal = itmCopier.SubItems(I_COL_SOUM_TOTAL)

195         m_arr_tyCopie(iIndex).sProfit = itmCopier.SubItems(I_COL_SOUM_PROFIT)

200         iIndex = iIndex + 1
205       End If
210     Next

215     Screen.MousePointer = vbDefault

220     Exit Sub

AfficherErreur:

225     woups "frmProjSoumElec", "CopierPiece", Err, Erl
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

100             Exit Sub
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

265       itmColler.SubItems(I_COL_SOUM_TEMPS) = m_arr_tyCopie(iCompteur).sTemps

270       itmColler.SubItems(I_COL_SOUM_MONTAGE) = m_arr_tyCopie(iCompteur).sMontage

275       itmColler.SubItems(I_COL_SOUM_TOTAL) = m_arr_tyCopie(iCompteur).sTotal

280       itmColler.SubItems(I_COL_SOUM_PROFIT) = m_arr_tyCopie(iCompteur).sProfit

285       If m_eType = TYPE_PROJET Then
290         itmColler.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = m_arr_tyCopie(iCompteur).lColor
295         itmColler.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = m_arr_tyCopie(iCompteur).lColor
300         itmColler.ListSubItems(I_COL_SOUM_MONTAGE).ForeColor = m_arr_tyCopie(iCompteur).lColor
305         itmColler.ListSubItems(I_COL_SOUM_TEMPS).ForeColor = m_arr_tyCopie(iCompteur).lColor
310         itmColler.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = m_arr_tyCopie(iCompteur).lColor
315         itmColler.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = m_arr_tyCopie(iCompteur).lColor
320         itmColler.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = m_arr_tyCopie(iCompteur).lColor
325         itmColler.ListSubItems(I_COL_SOUM_DESCR).ForeColor = m_arr_tyCopie(iCompteur).lColor
330         itmColler.ListSubItems(I_COL_SOUM_PIECE).ForeColor = m_arr_tyCopie(iCompteur).lColor
335         itmColler.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = m_arr_tyCopie(iCompteur).lColor
340         itmColler.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = m_arr_tyCopie(iCompteur).lColor
345       Else
350         itmColler.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = COLOR_NOIR
355         itmColler.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = COLOR_NOIR
360         itmColler.ListSubItems(I_COL_SOUM_MONTAGE).ForeColor = COLOR_NOIR
365         itmColler.ListSubItems(I_COL_SOUM_TEMPS).ForeColor = COLOR_NOIR
370         itmColler.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = COLOR_NOIR
375         itmColler.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = COLOR_NOIR
380         itmColler.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = COLOR_NOIR
385         itmColler.ListSubItems(I_COL_SOUM_DESCR).ForeColor = COLOR_NOIR
390         itmColler.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_NOIR
395         itmColler.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = COLOR_NOIR
400         itmColler.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = COLOR_NOIR
405       End If

410       Call lvwSoumission.Refresh
415     Next

420     Call CalculerTempsFabrication

425     Call CalculerPrix

430     Screen.MousePointer = vbDefault

435     Exit Sub

AfficherErreur:

440     woups "frmProjSoumElec", "CollerPiece", Err, Erl
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

45      woups "frmProjSoumElec", "Deselect", Err, Erl
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

60      woups "frmProjSoumElec", "txtPrixSpecial_Change", Err, Erl
End Sub

Private Sub txtPrixSpecial_LostFocus()

5       On Error GoTo AfficherErreur

10      txtPrixSpecial.Text = Replace(txtPrixSpecial.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmProjSoumElec", "txtPrixSpecial_LostFocus", Err, Erl
End Sub

Private Function ValiderFormatElectrique(ByVal sNoProjSoum As String) As Boolean
        
5       On Error GoTo AfficherErreur

10      If UCase(Left$(sNoProjSoum, 1)) = "E" Then
15        ValiderFormatElectrique = True
20      Else
25        Call MsgBox("Un numéro électrique doit absolument commencé par 'E' !", vbOKOnly, "Erreur")

30        ValiderFormatElectrique = False
35      End If

40      Exit Function

AfficherErreur:

45      woups "FrmProjSoumElec", "ValiderFormatElectrique", Err, Erl
End Function

Private Function ValiderFormatSoumission(ByVal sNoSoumission As String) As Boolean
        
5       On Error GoTo AfficherErreur

10      If Mid$(sNoSoumission, 3, 1) = "1" Then
15        ValiderFormatSoumission = True
20      Else
25        Call MsgBox("Une soumission doit absolument avoir un '1' comme 3e caractère !", vbOKOnly, "Erreur")

30        ValiderFormatSoumission = False
35      End If

40      Exit Function

AfficherErreur:

45      woups "FrmProjSoumElec", "ValiderFormatSoumission", Err, Erl
End Function

Private Function ValiderFormatJobSansSoum(ByVal sNoProjet As String) As Boolean
        
5       On Error GoTo AfficherErreur

10      If Mid$(sNoProjet, 3, 1) <> "3" And Mid$(sNoProjet, 3, 1) <> "1" Then
15        ValiderFormatJobSansSoum = True
20      Else
25        Call MsgBox("Un projet créé sans soumission ne peut pas être un '" & Mid$(sNoProjet, 2, 2) & "' !", vbOKOnly, "Erreur")

30        ValiderFormatJobSansSoum = False
35      End If

40      Exit Function

AfficherErreur:

45      woups "FrmProjSoumElec", "ValiderFormatJobSansSoum", Err, Erl
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

45      woups "FrmProjSoumElec", "ValiderFormatJobAvecSoum", Err, Erl
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

45      woups "FrmProjSoumElec", "ValiderFormatJobExtra", Err, Erl
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

100     Call rstProjCumulatif.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & sNoCumulatif & "'", g_connData, adOpenDynamic, adLockOptimistic)

105     If rstProjCumulatif.EOF Then
110       bCumulatifExiste = False

115       Call rstProjCumulatif.AddNew

120       rstProjCumulatif.Fields("IDProjet") = sNoCumulatif

          'Ouverture du projet -01 pour voir la soumission reliée pour ensuite assigner
          'la soumission -99 avec le projet -99
125       Call rstProj.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, 6) & "-01'", g_connData, adOpenForwardOnly, adLockReadOnly)

130       If Not rstProj.EOF Then
135         If Not IsNull(rstProj.Fields("IDSoumission")) Then
140           If Len(rstProj.Fields("IDSoumission")) >= 6 Then
145             Set rstSoum = New ADODB.Recordset

150             Call rstSoum.Open("SELECT * FROM GRB_SoumissionElec WHERE IDSoumission = '" & Left$(rstProj.Fields("IDSoumission"), 6) & "-99'", g_connData, adOpenForwardOnly, adLockReadOnly)

155             If Not rstSoum.EOF Then
160               rstProjCumulatif.Fields("IDSoumission") = rstSoum.Fields("IDSoumission")
165             End If

170             Call rstSoum.Close
175             Set rstSoum = Nothing
180           End If
185         End If
190       End If

195       Call rstProj.Close

200       Call rstProj.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

205       rstProjCumulatif.Fields("IDClient") = rstProj.Fields("IDClient")
210       rstProjCumulatif.Fields("IDContact") = rstProj.Fields("IDContact")

215       rstProjCumulatif.Fields("TauxDessin") = rstProj.Fields("TauxDessin")
220       rstProjCumulatif.Fields("TauxFabrication") = rstProj.Fields("TauxFabrication")
225       rstProjCumulatif.Fields("TauxAssemblage") = rstProj.Fields("TauxAssemblage")
230       rstProjCumulatif.Fields("TauxProgInterface") = rstProj.Fields("TauxProgInterface")
235       rstProjCumulatif.Fields("TauxProgAutomate") = rstProj.Fields("TauxProgAutomate")
240       rstProjCumulatif.Fields("TauxProgRobot") = rstProj.Fields("TauxProgRobot")
245       rstProjCumulatif.Fields("TauxVision") = rstProj.Fields("TauxVision")
250       rstProjCumulatif.Fields("TauxTest") = rstProj.Fields("TauxTest")
255       rstProjCumulatif.Fields("TauxInstallation") = rstProj.Fields("TauxInstallation")
260       rstProjCumulatif.Fields("TauxMiseService") = rstProj.Fields("TauxMiseService")
265       rstProjCumulatif.Fields("TauxFormation") = rstProj.Fields("TauxFormation")
270       rstProjCumulatif.Fields("TauxGestion") = rstProj.Fields("TauxGestion")
275       rstProjCumulatif.Fields("TauxShipping") = rstProj.Fields("TauxShipping")

280       rstProjCumulatif.Fields("Transport") = rstProj.Fields("Transport")

285       rstProjCumulatif.Fields("Profit") = rstProj.Fields("Profit")
290       rstProjCumulatif.Fields("imprevue") = rstProj.Fields("imprevue")
295       rstProjCumulatif.Fields("commission") = rstProj.Fields("commission")

300       Call rstProj.Close

305       rstProjCumulatif.Fields("Description") = "Cumulatif de " & Left$(txtNoProjSoum.Text, 6)

310       Set rstEmploye = New ADODB.Recordset

315       Call rstEmploye.Open("SELECT * FROM GRB_employés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

320       rstProjCumulatif.Fields("creer") = ConvertDate(Date)

325       rstProjCumulatif.Fields("creer_par") = rstEmploye.Fields("NoEmploye")

330       Call rstEmploye.Close
335       Set rstEmploye = Nothing

340       Call rstProjCumulatif.Update

345       Set rstProjSoum = New ADODB.Recordset

350       Call rstProjSoum.Open("SELECT * FROM GRB_ProjSoum", g_connData, adOpenDynamic, adLockOptimistic)

355       Call rstProjSoum.AddNew

360       rstProjSoum.Fields("IDProjSoum") = sNoCumulatif
365       rstProjSoum.Fields("NoClient") = rstProjCumulatif.Fields("IDClient")
370       rstProjSoum.Fields("Description") = rstProjCumulatif.Fields("Description")
375       rstProjSoum.Fields("DateOuverture") = ConvertDate(Date)
380       rstProjSoum.Fields("Ouvert") = True
385       rstProjSoum.Fields("Verrouillé") = True
390       rstProjSoum.Fields("Type") = "P"

395       Call rstProjSoum.Update
    
400       Call rstProjSoum.Close
405       Set rstProjSoum = Nothing
410     Else
415       bCumulatifExiste = True
420     End If

425     rstProj.CursorLocation = adUseClient

430     Call rstProj.Open("SELECT * FROM GRB_ProjetElec WHERE LEFT(IDProjet, 6) = '" & Left$(txtNoProjSoum.Text, 6) & "' AND RIGHT(IDProjet, 2) <> '99'", g_connData, adOpenForwardOnly, adLockReadOnly)

435     If rstProj.RecordCount = 1 Then
440       rstProjCumulatif.Fields("NbreManuel") = rstProj.Fields("NbreManuel")

445       rstProjCumulatif.Fields("PrixEmballage") = rstProj.Fields("PrixEmballage")

450       rstProjCumulatif.Fields("total_manuel") = rstProj.Fields("total_manuel")

455       rstProjCumulatif.Fields("MontantForfait") = rstProj.Fields("MontantForfait")
460     Else
465       Do While Not rstProj.EOF
470         If Not IsNull(rstProj.Fields("NbreManuel")) Then
475           dblNbreManuel = dblNbreManuel + CDbl(rstProj.Fields("NbreManuel"))
480         End If

485         If Not IsNull(rstProj.Fields("PrixEmballage")) Then
490           dblPrixEmballage = dblPrixEmballage + CDbl(rstProj.Fields("PrixEmballage"))
495         End If

500         If Not IsNull(rstProj.Fields("total_manuel")) Then
505           dblTotalManuel = dblTotalManuel + CDbl(rstProj.Fields("total_manuel"))
510         End If

515         If Not IsNull(rstProj.Fields("MontantForfait")) Then
520           If IsNumeric(rstProj.Fields("MontantForfait")) Then
525             dblForfait = dblForfait + CDbl(rstProj.Fields("MontantForfait"))
530           End If
535         End If

540         Call rstProj.MoveNext
545       Loop

550       rstProjCumulatif.Fields("NbreManuel") = dblNbreManuel
555       rstProjCumulatif.Fields("PrixEmballage") = dblPrixEmballage
560       rstProjCumulatif.Fields("total_manuel") = dblTotalManuel
565       rstProjCumulatif.Fields("MontantForfait") = dblForfait
570     End If

575     Call rstProj.Close

580     Call rstProjCumulatif.Update

585     Call rstProjCumulatif.Close

        'AJOUT DES PIÈCES
590     Call rstPiecesCumulatif.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoCumulatif & "'", g_connData, adOpenDynamic, adLockOptimistic)

595     If bCumulatifExiste = True Then
600       Call g_connData.Execute("DELETE * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoCumulatif & "' AND Provenance = '" & Right$(txtNoProjSoum.Text, 2) & "'")

605       Call rstPieces.Open("SELECT * FROM GRB_Projet_Pieces WHERE IDProjet = '" & txtNoProjSoum.Text & "' AND Provenance Is Null OR Provenance = '' ORDER BY NuméroLigne", g_connData, adOpenForwardOnly, adLockReadOnly)
610     Else
615       Call rstPieces.Open("SELECT * FROM GRB_Projet_Pieces WHERE LEFT(IDProjet, 6) = '" & Left$(txtNoProjSoum.Text, 6) & "' AND RIGHT(IDProjet, 2) <> '99' AND Provenance Is Null OR Provenance = '' ORDER BY CInt(OrdreSection), NuméroLigne", g_connData, adOpenForwardOnly, adLockReadOnly)
620     End If

625     Do While Not rstPieces.EOF
630       Call rstPiecesCumulatif.AddNew

635       rstPiecesCumulatif.Fields("IDProjet") = sNoCumulatif
640       rstPiecesCumulatif.Fields("IDSection") = rstPieces.Fields("IDSection")
645       rstPiecesCumulatif.Fields("NumItem") = rstPieces.Fields("NumItem")
650       rstPiecesCumulatif.Fields("Qté") = rstPieces.Fields("Qté")
655       rstPiecesCumulatif.Fields("Desc_FR") = rstPieces.Fields("Desc_FR")
660       rstPiecesCumulatif.Fields("Desc_EN") = rstPieces.Fields("Desc_EN")
665       rstPiecesCumulatif.Fields("Manufact") = rstPieces.Fields("Manufact")
670       rstPiecesCumulatif.Fields("Prix_list") = rstPieces.Fields("Prix_list")
675       rstPiecesCumulatif.Fields("Escompte") = rstPieces.Fields("Escompte")
680       rstPiecesCumulatif.Fields("Prix_net") = rstPieces.Fields("Prix_net")
685       rstPiecesCumulatif.Fields("IDFRS") = rstPieces.Fields("IDFRS")
690       rstPiecesCumulatif.Fields("Temps") = rstPieces.Fields("Temps")
695       rstPiecesCumulatif.Fields("Temps_Total") = rstPieces.Fields("Temps_Total")
700       rstPiecesCumulatif.Fields("Prix_total") = rstPieces.Fields("Prix_total")
705       rstPiecesCumulatif.Fields("Profit_Argent") = rstPieces.Fields("Profit_Argent")
710       rstPiecesCumulatif.Fields("SousSection") = rstPieces.Fields("SousSection")
715       rstPiecesCumulatif.Fields("OrdreSection") = rstPieces.Fields("OrdreSection")
720       rstPiecesCumulatif.Fields("NuméroLigne") = rstPieces.Fields("NuméroLigne")
725       rstPiecesCumulatif.Fields("PrixOrigine") = rstPieces.Fields("PrixOrigine")
730       rstPiecesCumulatif.Fields("Type") = rstPieces.Fields("Type")
735       rstPiecesCumulatif.Fields("Visible") = rstPieces.Fields("Visible")
740       rstPiecesCumulatif.Fields("Commentaire") = rstPieces.Fields("Commentaire")
745       rstPiecesCumulatif.Fields("Devise") = rstPieces.Fields("Devise")
750       rstPiecesCumulatif.Fields("Provenance") = Right(rstPieces.Fields("IDProjet"), 2)

755       Call rstPiecesCumulatif.Update

760       Call rstPieces.MoveNext
765     Loop

770     Call rstPiecesCumulatif.Close
775     Call rstPieces.Close

780     Set rstProj = Nothing
785     Set rstPieces = Nothing
790     Set rstProjCumulatif = Nothing
795     Set rstPiecesCumulatif = Nothing

800     Call CalculerTotalRecordset(sNoCumulatif)

805     If bCumulatifExiste = False Then
810       If cmbOuvertFerme.ListIndex = I_CMB_OUVERT Then
815         Call RemplirComboProjSoum(txtNoProjSoum.Text)
820       End If
825     End If

830     Exit Sub

AfficherErreur:

835     woups "FrmProjSoumElec", "AjouterProjetAuCumulatif", Err, Erl
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
55      Dim dblTempsDessin        As Double
60      Dim dblTempsFabrication   As Double
65      Dim dblTempsAssemblage    As Double
70      Dim dblTempsProgInterface As Double
75      Dim dblTempsProgAutomate  As Double
80      Dim dblTempsProgRobot     As Double
85      Dim dblTempsVision        As Double
90      Dim dblTempsTest          As Double
95      Dim dblTempsInstallation  As Double
100     Dim dblTempsMiseService   As Double
105     Dim dblTempsFormation     As Double
110     Dim dblTempsGestion       As Double
115     Dim dblTempsShipping      As Double
120     Dim dblTempsTransport     As Double
125     Dim dblTempsUniteMobile   As Double
130     Dim dblTotalHebergement   As Double
135     Dim dblTotalRepas         As Double
140     Dim dblPrixEmballage      As Double
145     Dim dblTotalManuel        As Double
150     Dim dblForfait            As Double

155     sNoCumulatif = Left$(txtNoProjSoum.Text, 7) & "99"

160     Set rstSoum = New ADODB.Recordset
165     Set rstPieces = New ADODB.Recordset
170     Set rstSoumCumulatif = New ADODB.Recordset
175     Set rstPiecesCumulatif = New ADODB.Recordset

180     Call rstSoumCumulatif.Open("SELECT * FROM GRB_SoumissionElec WHERE IDSoumission = '" & sNoCumulatif & "'", g_connData, adOpenDynamic, adLockOptimistic)

185     If rstSoumCumulatif.EOF Then
190       bCumulatifExiste = False

195       Call rstSoumCumulatif.AddNew

200       rstSoumCumulatif.Fields("IDSoumission") = sNoCumulatif

205       Call rstSoum.Open("SELECT * FROM GRB_SoumissionElec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

210       rstSoumCumulatif.Fields("IDClient") = rstSoum.Fields("IDClient")
215       rstSoumCumulatif.Fields("IDContact") = rstSoum.Fields("IDContact")

220       rstSoumCumulatif.Fields("TauxDessin") = rstSoum.Fields("TauxDessin")
225       rstSoumCumulatif.Fields("TauxFabrication") = rstSoum.Fields("TauxFabrication")
230       rstSoumCumulatif.Fields("TauxAssemblage") = rstSoum.Fields("TauxAssemblage")
235       rstSoumCumulatif.Fields("TauxProgInterface") = rstSoum.Fields("TauxProgInterface")
240       rstSoumCumulatif.Fields("TauxProgAutomate") = rstSoum.Fields("TauxProgAutomate")
245       rstSoumCumulatif.Fields("TauxProgRobot") = rstSoum.Fields("TauxProgRobot")
250       rstSoumCumulatif.Fields("TauxVision") = rstSoum.Fields("TauxVision")
255       rstSoumCumulatif.Fields("TauxTest") = rstSoum.Fields("TauxTest")
260       rstSoumCumulatif.Fields("TauxInstallation") = rstSoum.Fields("TauxInstallation")
265       rstSoumCumulatif.Fields("TauxMiseService") = rstSoum.Fields("TauxMiseService")
270       rstSoumCumulatif.Fields("TauxFormation") = rstSoum.Fields("TauxFormation")
275       rstSoumCumulatif.Fields("TauxGestion") = rstSoum.Fields("TauxGestion")
280       rstSoumCumulatif.Fields("TauxShipping") = rstSoum.Fields("TauxShipping")

285       rstSoumCumulatif.Fields("TauxHebergement1") = rstSoum.Fields("TauxHebergement1")
290       rstSoumCumulatif.Fields("TauxHebergement2") = rstSoum.Fields("TauxHebergement2")
295       rstSoumCumulatif.Fields("TauxRepas") = rstSoum.Fields("TauxRepas")
300       rstSoumCumulatif.Fields("TauxTransport") = rstSoum.Fields("TauxTransport")
305       rstSoumCumulatif.Fields("TauxUniteMobile") = rstSoum.Fields("TauxUniteMobile")

310       rstSoumCumulatif.Fields("Transport") = rstSoum.Fields("Transport")

315       rstSoumCumulatif.Fields("Profit") = rstSoum.Fields("Profit")
320       rstSoumCumulatif.Fields("imprevue") = rstSoum.Fields("imprevue")
325       rstSoumCumulatif.Fields("commission") = rstSoum.Fields("commission")

330       Call rstSoum.Close

335       rstSoumCumulatif.Fields("Description") = "Cumulatif de " & Left$(txtNoProjSoum.Text, 6)

340       Set rstEmploye = New ADODB.Recordset

345       Call rstEmploye.Open("SELECT * FROM GRB_employés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

350       rstSoumCumulatif.Fields("creer") = ConvertDate(Date)

355       rstSoumCumulatif.Fields("creer_par") = rstEmploye.Fields("NoEmploye")

360       Call rstEmploye.Close
365       Set rstEmploye = Nothing

370       Call rstSoumCumulatif.Update

375       Set rstProjSoum = New ADODB.Recordset

380       Call rstProjSoum.Open("SELECT * FROM GRB_ProjSoum", g_connData, adOpenDynamic, adLockOptimistic)

385       Call rstProjSoum.AddNew

390       rstProjSoum.Fields("IDProjSoum") = sNoCumulatif
395       rstProjSoum.Fields("NoClient") = rstSoumCumulatif.Fields("IDClient")
400       rstProjSoum.Fields("Description") = rstSoumCumulatif.Fields("Description")
405       rstProjSoum.Fields("DateOuverture") = ConvertDate(Date)
410       rstProjSoum.Fields("Ouvert") = True
415       rstProjSoum.Fields("Verrouillé") = True
420       rstProjSoum.Fields("Type") = "S"

425       Call rstProjSoum.Update
    
430       Call rstProjSoum.Close
435       Set rstProjSoum = Nothing
440     Else
445       bCumulatifExiste = True
450     End If
     
455     rstSoum.CursorLocation = adUseClient
     
460     Call rstSoum.Open("SELECT * FROM GRB_SoumissionElec WHERE LEFT(IDSoumission, 6) = '" & Left$(txtNoProjSoum.Text, 6) & "' AND RIGHT(IDSoumission, 2) <> '99'", g_connData, adOpenForwardOnly, adLockReadOnly)

465     If rstSoum.RecordCount = 1 Then
470       rstSoumCumulatif.Fields("NbreManuel") = rstSoum.Fields("NbreManuel")

475       rstSoumCumulatif.Fields("TempsDessin") = rstSoum.Fields("TempsDessin")

480       If rstSoum.Fields("SansTemps") = False Then
485         rstSoumCumulatif.Fields("TempsFabrication") = rstSoum.Fields("TempsFabrication")
490       Else
495         rstSoumCumulatif.Fields("TempsFabrication") = 0
500       End If

505       rstSoumCumulatif.Fields("TempsAssemblage") = rstSoum.Fields("TempsAssemblage")
510       rstSoumCumulatif.Fields("TempsProgInterface") = rstSoum.Fields("TempsProgInterface")
515       rstSoumCumulatif.Fields("TempsProgAutomate") = rstSoum.Fields("TempsProgAutomate")
520       rstSoumCumulatif.Fields("TempsProgRobot") = rstSoum.Fields("TempsProgRobot")
525       rstSoumCumulatif.Fields("TempsVision") = rstSoum.Fields("TempsVision")
530       rstSoumCumulatif.Fields("TempsTest") = rstSoum.Fields("TempsTest")
535       rstSoumCumulatif.Fields("TempsInstallation") = rstSoum.Fields("TempsInstallation")
540       rstSoumCumulatif.Fields("TempsMiseService") = rstSoum.Fields("TempsMiseService")
545       rstSoumCumulatif.Fields("TempsFormation") = rstSoum.Fields("TempsFormation")
550       rstSoumCumulatif.Fields("TempsGestion") = rstSoum.Fields("TempsGestion")
555       rstSoumCumulatif.Fields("TempsShipping") = rstSoum.Fields("TempsShipping")

560       rstSoumCumulatif.Fields("NbrePersonne") = rstSoum.Fields("NbrePersonne")
565       rstSoumCumulatif.Fields("TempsHebergement") = rstSoum.Fields("TempsHebergement")
570       rstSoumCumulatif.Fields("TempsRepas") = rstSoum.Fields("TempsRepas")
575       rstSoumCumulatif.Fields("TempsTransport") = rstSoum.Fields("TempsTransport")
580       rstSoumCumulatif.Fields("TempsUniteMobile") = rstSoum.Fields("TempsUniteMobile")

585       rstSoumCumulatif.Fields("TotalHebergement") = rstSoum.Fields("TotalHebergement")
590       rstSoumCumulatif.Fields("TotalRepas") = rstSoum.Fields("TotalRepas")
595       rstSoumCumulatif.Fields("PrixEmballage") = rstSoum.Fields("PrixEmballage")

600       rstSoumCumulatif.Fields("total_manuel") = rstSoum.Fields("total_manuel")

605       rstSoumCumulatif.Fields("MontantForfait") = rstSoum.Fields("MontantForfait")
610     Else
615       Do While Not rstSoum.EOF
620         If Not IsNull(rstSoum.Fields("NbreManuel")) Then
625           dblNbreManuel = dblNbreManuel + CDbl(rstSoum.Fields("NbreManuel"))
630         End If

635         If Not IsNull(rstSoum.Fields("TempsDessin")) Then
640           dblTempsDessin = dblTempsDessin + CDbl(rstSoum.Fields("TempsDessin"))
645         End If

650         If rstSoum.Fields("SansTemps") = False Then
655           If Not IsNull(rstSoum.Fields("TempsFabrication")) Then
660             dblTempsFabrication = dblTempsFabrication + CDbl(rstSoum.Fields("TempsFabrication"))
665           End If
670         End If

675         If Not IsNull(rstSoum.Fields("TempsAssemblage")) Then
680           dblTempsAssemblage = dblTempsAssemblage + CDbl(rstSoum.Fields("TempsAssemblage"))
685         End If

690         If Not IsNull(rstSoum.Fields("TempsProgInterface")) Then
695           dblTempsProgInterface = dblTempsProgInterface + CDbl(rstSoum.Fields("TempsProgInterface"))
700         End If

705         If Not IsNull(rstSoum.Fields("TempsProgAutomate")) Then
710           dblTempsProgAutomate = dblTempsProgAutomate + CDbl(rstSoum.Fields("TempsProgAutomate"))
715         End If

720         If Not IsNull(rstSoum.Fields("TempsProgRobot")) Then
725           dblTempsProgRobot = dblTempsProgRobot + CDbl(rstSoum.Fields("TempsProgRobot"))
730         End If

735         If Not IsNull(rstSoum.Fields("TempsVision")) Then
740           dblTempsVision = dblTempsVision + CDbl(rstSoum.Fields("TempsVision"))
745         End If

750         If Not IsNull(rstSoum.Fields("TempsTest")) Then
755           dblTempsTest = dblTempsTest + CDbl(rstSoum.Fields("TempsTest"))
760         End If

765         If Not IsNull(rstSoum.Fields("TempsInstallation")) Then
770           dblTempsInstallation = dblTempsInstallation + CDbl(rstSoum.Fields("TempsInstallation"))
775         End If

780         If Not IsNull(rstSoum.Fields("TempsMiseService")) Then
785           dblTempsMiseService = dblTempsMiseService + CDbl(rstSoum.Fields("TempsMiseService"))
790         End If

795         If Not IsNull(rstSoum.Fields("TempsFormation")) Then
800           dblTempsFormation = dblTempsFormation + CDbl(rstSoum.Fields("TempsFormation"))
805         End If

810         If Not IsNull(rstSoum.Fields("TempsGestion")) Then
815           dblTempsGestion = dblTempsGestion + CDbl(rstSoum.Fields("TempsGestion"))
820         End If

825         If Not IsNull(rstSoum.Fields("TempsShipping")) Then
830           dblTempsShipping = dblTempsShipping + CDbl(rstSoum.Fields("TempsShipping"))
835         End If

840         If Not IsNull(rstSoum.Fields("TempsTransport")) Then
845           dblTempsTransport = dblTempsTransport + CDbl(rstSoum.Fields("TempsTransport"))
850         End If

855         If Not IsNull(rstSoum.Fields("TempsUniteMobile")) Then
860           dblTempsUniteMobile = dblTempsUniteMobile + CDbl(rstSoum.Fields("TempsUniteMobile"))
865         End If

870         If Not IsNull(rstSoum.Fields("TotalHebergement")) Then
875           dblTotalHebergement = dblTotalHebergement + CDbl(rstSoum.Fields("TotalHebergement"))
880         End If

885         If Not IsNull(rstSoum.Fields("TotalRepas")) Then
890           dblTotalRepas = dblTotalRepas + CDbl(rstSoum.Fields("TotalRepas"))
895         End If

900         If Not IsNull(rstSoum.Fields("PrixEmballage")) Then
905           dblPrixEmballage = dblPrixEmballage + CDbl(rstSoum.Fields("PrixEmballage"))
910         End If

915         If Not IsNull(rstSoum.Fields("total_manuel")) Then
920           dblTotalManuel = dblTotalManuel + CDbl(rstSoum.Fields("total_manuel"))
925         End If

930         If Not IsNull(rstSoum.Fields("MontantForfait")) Then
935           If IsNumeric(rstSoum.Fields("MontantForfait")) Then
940             dblForfait = dblForfait + CDbl(rstSoum.Fields("MontantForfait"))
945           End If
950         End If

955         Call rstSoum.MoveNext
960       Loop

965       rstSoumCumulatif.Fields("NbreManuel") = dblNbreManuel

970       rstSoumCumulatif.Fields("TempsDessin") = dblTempsDessin
975       rstSoumCumulatif.Fields("TempsFabrication") = dblTempsFabrication
980       rstSoumCumulatif.Fields("TempsAssemblage") = dblTempsAssemblage
985       rstSoumCumulatif.Fields("TempsProgInterface") = dblTempsProgInterface
990       rstSoumCumulatif.Fields("TempsProgAutomate") = dblTempsProgAutomate
995       rstSoumCumulatif.Fields("TempsProgRobot") = dblTempsProgRobot
1000      rstSoumCumulatif.Fields("TempsVision") = dblTempsVision
1005      rstSoumCumulatif.Fields("TempsTest") = dblTempsTest
1010      rstSoumCumulatif.Fields("TempsInstallation") = dblTempsInstallation
1015      rstSoumCumulatif.Fields("TempsMiseService") = dblTempsMiseService
1020      rstSoumCumulatif.Fields("TempsFormation") = dblTempsFormation
1025      rstSoumCumulatif.Fields("TempsGestion") = dblTempsGestion
1030      rstSoumCumulatif.Fields("TempsShipping") = dblTempsShipping

1035      rstSoumCumulatif.Fields("TempsTransport") = dblTempsTransport
1040      rstSoumCumulatif.Fields("TempsUniteMobile") = dblTempsUniteMobile

1045      rstSoumCumulatif.Fields("TotalHebergement") = dblTotalHebergement
1050      rstSoumCumulatif.Fields("TotalRepas") = dblTotalRepas
1055      rstSoumCumulatif.Fields("PrixEmballage") = dblPrixEmballage

1060      rstSoumCumulatif.Fields("total_manuel") = dblTotalManuel

1065      rstSoumCumulatif.Fields("MontantForfait") = dblForfait
1070    End If

1075    Call rstSoumCumulatif.Update

1080    Call rstSoumCumulatif.Close

1085    Call rstSoum.Close

        'AJOUT DES PIÈCES
1090    Call rstPiecesCumulatif.Open("SELECT * FROM GRB_Soumission_Pieces WHERE IDSoumission = '" & sNoCumulatif & "'", g_connData, adOpenDynamic, adLockOptimistic)
             
1095    If bCumulatifExiste = True Then
1100      Call g_connData.Execute("DELETE * FROM GRB_Soumission_Pieces WHERE IDSoumission = '" & sNoCumulatif & "' AND Provenance = '" & Right$(txtNoProjSoum.Text, 2) & "'")

1105      Call rstPieces.Open("SELECT * FROM GRB_Soumission_Pieces WHERE IDSoumission = '" & txtNoProjSoum.Text & "' ORDER BY NuméroLigne", g_connData, adOpenForwardOnly, adLockReadOnly)
1110    Else
1115      Call rstPieces.Open("SELECT * FROM GRB_Soumission_Pieces WHERE LEFT(IDSoumission, 6) = '" & Left$(txtNoProjSoum.Text, 6) & "' AND RIGHT(IDSoumission, 2) <> '99' ORDER BY CInt(OrdreSection), NuméroLigne", g_connData, adOpenForwardOnly, adLockReadOnly)
1120    End If

1125    Do While Not rstPieces.EOF
1130      Call rstPiecesCumulatif.AddNew

1135      rstPiecesCumulatif.Fields("IDSoumission") = sNoCumulatif
1140      rstPiecesCumulatif.Fields("IDSection") = rstPieces.Fields("IDSection")
1145      rstPiecesCumulatif.Fields("NumItem") = rstPieces.Fields("NumItem")
1150      rstPiecesCumulatif.Fields("Qté") = rstPieces.Fields("Qté")
1155      rstPiecesCumulatif.Fields("Desc_FR") = rstPieces.Fields("Desc_FR")
1160      rstPiecesCumulatif.Fields("Desc_EN") = rstPieces.Fields("Desc_EN")
1165      rstPiecesCumulatif.Fields("Manufact") = rstPieces.Fields("Manufact")
1170      rstPiecesCumulatif.Fields("Prix_list") = rstPieces.Fields("Prix_list")
1175      rstPiecesCumulatif.Fields("Escompte") = rstPieces.Fields("Escompte")
1180      rstPiecesCumulatif.Fields("Prix_net") = rstPieces.Fields("Prix_net")
1185      rstPiecesCumulatif.Fields("IDFRS") = rstPieces.Fields("IDFRS")
1190      rstPiecesCumulatif.Fields("Temps") = rstPieces.Fields("Temps")
1195      rstPiecesCumulatif.Fields("Temps_Total") = rstPieces.Fields("Temps_Total")
1200      rstPiecesCumulatif.Fields("Prix_total") = rstPieces.Fields("Prix_total")
1205      rstPiecesCumulatif.Fields("Profit_Argent") = rstPieces.Fields("Profit_Argent")
1210      rstPiecesCumulatif.Fields("SousSection") = rstPieces.Fields("SousSection")
1215      rstPiecesCumulatif.Fields("OrdreSection") = rstPieces.Fields("OrdreSection")
1220      rstPiecesCumulatif.Fields("NuméroLigne") = rstPieces.Fields("NuméroLigne")
1225      rstPiecesCumulatif.Fields("PrixOrigine") = rstPieces.Fields("PrixOrigine")
1230      rstPiecesCumulatif.Fields("Type") = rstPieces.Fields("Type")
1235      rstPiecesCumulatif.Fields("Visible") = rstPieces.Fields("Visible")
1240      rstPiecesCumulatif.Fields("Commentaire") = rstPieces.Fields("Commentaire")
1245      rstPiecesCumulatif.Fields("Devise") = rstPieces.Fields("Devise")
1250      rstPiecesCumulatif.Fields("Provenance") = Right(rstPieces.Fields("IDSoumission"), 2)

1255      Call rstPiecesCumulatif.Update

1260      Call rstPieces.MoveNext
1265    Loop

1270    Call rstPiecesCumulatif.Close
1275    Call rstPieces.Close

1280    Set rstSoum = Nothing
1285    Set rstPieces = Nothing
1290    Set rstSoumCumulatif = Nothing
1295    Set rstPiecesCumulatif = Nothing

1300    Call CalculerTotalRecordset(sNoCumulatif)

1305    Exit Sub

AfficherErreur:

1310    woups "FrmProjSoumElec", "AjouterSoumissionAuCumulatif", Err, Erl
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

95      Call rstProj.Open("SELECT * FROM GRB_ProjetElec WHERE LEFT(IDProjet, 6) = '" & Left$(txtNoProjSoum.Text, 6) & "' AND RIGHT(IDProjet, 2) <> '99'", g_connData, adOpenForwardOnly, adLockReadOnly)

100     If rstProj.EOF Then
105       Call g_connData.Execute("DELETE * FROM GRB_ProjSoum WHERE IDProjSoum = '" & sNoCumulatif & "'")

110       Call g_connData.Execute("DELETE * FROM GRB_Projet_Pieces WHERE IDProjet = '" & sNoCumulatif & "' AND Type = 'E'")

          'Efface le projet
115       Call g_connData.Execute("DELETE * FROM GRB_ProjetElec WHERE IDProjet = '" & sNoCumulatif & "'")

120       bSupprimer = True
125     Else
130       Call rstProjCumulatif.Open("SELECT * FROM GRB_ProjetElec WHERE IDProjet = '" & sNoCumulatif & "'", g_connData, adOpenDynamic, adLockOptimistic)

135       If rstProj.RecordCount = 1 Then
140         rstProjCumulatif.Fields("NbreManuel") = rstProj.Fields("NbreManuel")

145         rstProjCumulatif.Fields("PrixEmballage") = rstProj.Fields("PrixEmballage")

150         rstProjCumulatif.Fields("total_manuel") = rstProj.Fields("total_manuel")

155         rstProjCumulatif.Fields("MontantForfait") = rstProj.Fields("MontantForfait")
160       Else
165         Do While Not rstProj.EOF
170           If Not IsNull(rstProj.Fields("NbreManuel")) Then
175             dblNbreManuel = dblNbreManuel + CDbl(rstProj.Fields("NbreManuel"))
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

250         rstProjCumulatif.Fields("NbreManuel") = dblNbreManuel
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
375       rstPiecesCumulatif.Fields("Temps") = rstPieces.Fields("Temps")
380       rstPiecesCumulatif.Fields("Temps_Total") = rstPieces.Fields("Temps_Total")
385       rstPiecesCumulatif.Fields("Prix_total") = rstPieces.Fields("Prix_total")
390       rstPiecesCumulatif.Fields("Profit_Argent") = rstPieces.Fields("Profit_Argent")
395       rstPiecesCumulatif.Fields("SousSection") = rstPieces.Fields("SousSection")
400       rstPiecesCumulatif.Fields("OrdreSection") = rstPieces.Fields("OrdreSection")
405       rstPiecesCumulatif.Fields("NuméroLigne") = rstPieces.Fields("NuméroLigne")
410       rstPiecesCumulatif.Fields("PrixOrigine") = rstPieces.Fields("PrixOrigine")
415       rstPiecesCumulatif.Fields("Type") = rstPieces.Fields("Type")
420       rstPiecesCumulatif.Fields("Visible") = rstPieces.Fields("Visible")
425       rstPiecesCumulatif.Fields("Commentaire") = rstPieces.Fields("Commentaire")
430       rstPiecesCumulatif.Fields("Devise") = rstPieces.Fields("Devise")
435       rstPiecesCumulatif.Fields("Provenance") = Right(rstPieces.Fields("IDProjet"), 2)

440       Call rstPiecesCumulatif.Update

445       Call rstPieces.MoveNext
450     Loop

455     Call rstPiecesCumulatif.Close
460     Call rstPieces.Close

465     Set rstProj = Nothing
470     Set rstPieces = Nothing
475     Set rstProjCumulatif = Nothing
480     Set rstPiecesCumulatif = Nothing

485     If bSupprimer = False Then
490       Call CalculerTotalRecordset(sNoCumulatif)
495     End If

500     Exit Sub

AfficherErreur:

505     woups "FrmProjSoumElec", "AjouterProjetAuCumulatif", Err, Erl
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
45      Dim dblTempsDessin        As Double
50      Dim dblTempsFabrication   As Double
55      Dim dblTempsAssemblage    As Double
60      Dim dblTempsProgInterface As Double
65      Dim dblTempsProgAutomate  As Double
70      Dim dblTempsProgRobot     As Double
75      Dim dblTempsVision        As Double
80      Dim dblTempsTest          As Double
85      Dim dblTempsInstallation  As Double
90      Dim dblTempsMiseService   As Double
95      Dim dblTempsFormation     As Double
100     Dim dblTempsGestion       As Double
105     Dim dblTempsShipping      As Double
110     Dim dblTempsTransport     As Double
115     Dim dblTempsUniteMobile   As Double
120     Dim dblTotalHebergement   As Double
125     Dim dblTotalRepas         As Double
130     Dim dblPrixEmballage      As Double
135     Dim dblTotalManuel        As Double
140     Dim dblForfait            As Double
145     Dim bSupprimer            As Boolean

150     sNoCumulatif = Left$(txtNoProjSoum.Text, 7) & "99"

155     Set rstSoum = New ADODB.Recordset
160     Set rstPieces = New ADODB.Recordset
165     Set rstSoumCumulatif = New ADODB.Recordset
170     Set rstPiecesCumulatif = New ADODB.Recordset
     
175     rstSoum.CursorLocation = adUseClient
     
180     Call rstSoum.Open("SELECT * FROM GRB_SoumissionElec WHERE LEFT(IDSoumission, 6) = '" & Left$(txtNoProjSoum.Text, 6) & "' AND RIGHT(IDSoumission, 2) <> '99'", g_connData, adOpenForwardOnly, adLockReadOnly)

185     If rstSoum.EOF Then
190       Call g_connData.Execute("DELETE * FROM GRB_ProjSoum WHERE IDProjSoum = '" & sNoCumulatif & "'")

195       Call g_connData.Execute("DELETE * FROM GRB_Soumission_Pieces WHERE IDSoumission = '" & sNoCumulatif & "' AND Type = 'E'")
                
          'Efface la soumission
200       Call g_connData.Execute("DELETE * FROM GRB_SoumissionElec WHERE IDSoumission = '" & sNoCumulatif & "'")

205       bSupprimer = True
210     Else
215       Call rstSoumCumulatif.Open("SELECT * FROM GRB_SoumissionElec WHERE IDSoumission = '" & sNoCumulatif & "'", g_connData, adOpenDynamic, adLockOptimistic)

220       If rstSoum.RecordCount = 1 Then
225         rstSoumCumulatif.Fields("NbreManuel") = rstSoum.Fields("NbreManuel")
  
230         rstSoumCumulatif.Fields("TempsDessin") = rstSoum.Fields("TempsDessin")

235         If rstSoum.Fields("SansTemps") = False Then
240           rstSoumCumulatif.Fields("TempsFabrication") = rstSoum.Fields("TempsFabrication")
245         Else
250           rstSoumCumulatif.Fields("TempsFabrication") = 0
255         End If

260         rstSoumCumulatif.Fields("TempsAssemblage") = rstSoum.Fields("TempsAssemblage")
265         rstSoumCumulatif.Fields("TempsProgInterface") = rstSoum.Fields("TempsProgInterface")
270         rstSoumCumulatif.Fields("TempsProgAutomate") = rstSoum.Fields("TempsProgAutomate")
275         rstSoumCumulatif.Fields("TempsProgRobot") = rstSoum.Fields("TempsProgRobot")
280         rstSoumCumulatif.Fields("TempsVision") = rstSoum.Fields("TempsVision")
285         rstSoumCumulatif.Fields("TempsTest") = rstSoum.Fields("TempsTest")
290         rstSoumCumulatif.Fields("TempsInstallation") = rstSoum.Fields("TempsInstallation")
295         rstSoumCumulatif.Fields("TempsMiseService") = rstSoum.Fields("TempsMiseService")
300         rstSoumCumulatif.Fields("TempsFormation") = rstSoum.Fields("TempsFormation")
305         rstSoumCumulatif.Fields("TempsGestion") = rstSoum.Fields("TempsGestion")
310         rstSoumCumulatif.Fields("TempsShipping") = rstSoum.Fields("TempsShipping")
  
315         rstSoumCumulatif.Fields("NbrePersonne") = rstSoum.Fields("NbrePersonne")
320         rstSoumCumulatif.Fields("TempsHebergement") = rstSoum.Fields("TempsHebergement")
325         rstSoumCumulatif.Fields("TempsRepas") = rstSoum.Fields("TempsRepas")
330         rstSoumCumulatif.Fields("TempsTransport") = rstSoum.Fields("TempsTransport")
335         rstSoumCumulatif.Fields("TempsUniteMobile") = rstSoum.Fields("TempsUniteMobile")
  
340         rstSoumCumulatif.Fields("TotalHebergement") = rstSoum.Fields("TotalHebergement")
345         rstSoumCumulatif.Fields("TotalRepas") = rstSoum.Fields("TotalRepas")
350         rstSoumCumulatif.Fields("PrixEmballage") = rstSoum.Fields("PrixEmballage")

355         rstSoumCumulatif.Fields("total_manuel") = rstSoum.Fields("total_manuel")

360         rstSoumCumulatif.Fields("MontantForfait") = rstSoum.Fields("MontantForfait")
365       Else
370         Do While Not rstSoum.EOF
375           If Not IsNull(rstSoum.Fields("NbreManuel")) Then
380             dblNbreManuel = dblNbreManuel + CDbl(rstSoum.Fields("NbreManuel"))
385           End If

390           If Not IsNull(rstSoum.Fields("TempsDessin")) Then
395             dblTempsDessin = dblTempsDessin + CDbl(rstSoum.Fields("TempsDessin"))
400           End If
  
405           If rstSoum.Fields("SansTemps") = False Then
410             If Not IsNull(rstSoum.Fields("TempsFabrication")) Then
415               dblTempsFabrication = dblTempsFabrication + CDbl(rstSoum.Fields("TempsFabrication"))
420             End If
425           End If
  
430           If Not IsNull(rstSoum.Fields("TempsAssemblage")) Then
435             dblTempsAssemblage = dblTempsAssemblage + CDbl(rstSoum.Fields("TempsAssemblage"))
440           End If
  
445           If Not IsNull(rstSoum.Fields("TempsProgInterface")) Then
450             dblTempsProgInterface = dblTempsProgInterface + CDbl(rstSoum.Fields("TempsProgInterface"))
455           End If
  
460           If Not IsNull(rstSoum.Fields("TempsProgAutomate")) Then
465             dblTempsProgAutomate = dblTempsProgAutomate + CDbl(rstSoum.Fields("TempsProgAutomate"))
470           End If
  
475           If Not IsNull(rstSoum.Fields("TempsProgRobot")) Then
480             dblTempsProgRobot = dblTempsProgRobot + CDbl(rstSoum.Fields("TempsProgRobot"))
485           End If
  
490           If Not IsNull(rstSoum.Fields("TempsVision")) Then
495             dblTempsVision = dblTempsVision + CDbl(rstSoum.Fields("TempsVision"))
500           End If
  
505           If Not IsNull(rstSoum.Fields("TempsTest")) Then
510             dblTempsTest = dblTempsTest + CDbl(rstSoum.Fields("TempsTest"))
515           End If
  
520           If Not IsNull(rstSoum.Fields("TempsInstallation")) Then
525             dblTempsInstallation = dblTempsInstallation + CDbl(rstSoum.Fields("TempsInstallation"))
530           End If
  
535           If Not IsNull(rstSoum.Fields("TempsMiseService")) Then
540             dblTempsMiseService = dblTempsMiseService + CDbl(rstSoum.Fields("TempsMiseService"))
545           End If
  
550           If Not IsNull(rstSoum.Fields("TempsFormation")) Then
555             dblTempsFormation = dblTempsFormation + CDbl(rstSoum.Fields("TempsFormation"))
560           End If
  
565           If Not IsNull(rstSoum.Fields("TempsGestion")) Then
570             dblTempsGestion = dblTempsGestion + CDbl(rstSoum.Fields("TempsGestion"))
575           End If
  
580           If Not IsNull(rstSoum.Fields("TempsShipping")) Then
585             dblTempsShipping = dblTempsShipping + CDbl(rstSoum.Fields("TempsShipping"))
590           End If
  
595           If Not IsNull(rstSoum.Fields("TempsTransport")) Then
600             dblTempsTransport = dblTempsTransport + CDbl(rstSoum.Fields("TempsTransport"))
605           End If
  
610           If Not IsNull(rstSoum.Fields("TempsUniteMobile")) Then
615             dblTempsUniteMobile = dblTempsUniteMobile + CDbl(rstSoum.Fields("TempsUniteMobile"))
620           End If
  
625           If Not IsNull(rstSoum.Fields("TotalHebergement")) Then
630             dblTotalHebergement = dblTotalHebergement + CDbl(rstSoum.Fields("TotalHebergement"))
635           End If
  
640           If Not IsNull(rstSoum.Fields("TotalRepas")) Then
645             dblTotalRepas = dblTotalRepas + CDbl(rstSoum.Fields("TotalRepas"))
650           End If
  
655           If Not IsNull(rstSoum.Fields("PrixEmballage")) Then
660             dblPrixEmballage = dblPrixEmballage + CDbl(rstSoum.Fields("PrixEmballage"))
665           End If
  
670           If Not IsNull(rstSoum.Fields("total_manuel")) Then
675             dblTotalManuel = dblTotalManuel + CDbl(rstSoum.Fields("total_manuel"))
680           End If

685           If Not IsNull(rstSoum.Fields("MontantForfait")) Then
690             If IsNumeric(rstSoum.Fields("MontantForfait")) Then
695               dblForfait = dblForfait + CDbl(rstSoum.Fields("MontantForfait"))
700             End If
705           End If
  
710           Call rstSoum.MoveNext
715         Loop
  
720         rstSoumCumulatif.Fields("NbreManuel") = dblNbreManuel
  
725         rstSoumCumulatif.Fields("TempsDessin") = dblTempsDessin
730         rstSoumCumulatif.Fields("TempsFabrication") = dblTempsFabrication
735         rstSoumCumulatif.Fields("TempsAssemblage") = dblTempsAssemblage
740         rstSoumCumulatif.Fields("TempsProgInterface") = dblTempsProgInterface
745         rstSoumCumulatif.Fields("TempsProgAutomate") = dblTempsProgAutomate
750         rstSoumCumulatif.Fields("TempsProgRobot") = dblTempsProgRobot
755         rstSoumCumulatif.Fields("TempsVision") = dblTempsVision
760         rstSoumCumulatif.Fields("TempsTest") = dblTempsTest
765         rstSoumCumulatif.Fields("TempsInstallation") = dblTempsInstallation
770         rstSoumCumulatif.Fields("TempsMiseService") = dblTempsMiseService
775         rstSoumCumulatif.Fields("TempsFormation") = dblTempsFormation
780         rstSoumCumulatif.Fields("TempsGestion") = dblTempsGestion
785         rstSoumCumulatif.Fields("TempsShipping") = dblTempsShipping
  
790         rstSoumCumulatif.Fields("TempsTransport") = dblTempsTransport
795         rstSoumCumulatif.Fields("TempsUniteMobile") = dblTempsUniteMobile
  
800         rstSoumCumulatif.Fields("TotalHebergement") = dblTotalHebergement
805         rstSoumCumulatif.Fields("TotalRepas") = dblTotalRepas
810         rstSoumCumulatif.Fields("PrixEmballage") = dblPrixEmballage
  
815         rstSoumCumulatif.Fields("total_manuel") = dblTotalManuel

820         rstSoumCumulatif.Fields("MontantForfait") = dblForfait
825       End If

830       Call rstSoum.Close

835       Call rstSoumCumulatif.Update

840       Call rstSoumCumulatif.Close
845     End If

        'AJOUT DES PIÈCES
850     Call rstPiecesCumulatif.Open("SELECT * FROM GRB_Soumission_Pieces WHERE IDSoumission = '" & sNoCumulatif & "'", g_connData, adOpenDynamic, adLockOptimistic)
             
855     Call g_connData.Execute("DELETE * FROM GRB_Soumission_Pieces WHERE IDSoumission = '" & sNoCumulatif & "'")

860     Call rstPieces.Open("SELECT * FROM GRB_Soumission_Pieces WHERE LEFT(IDSoumission, 6) = '" & Left$(txtNoProjSoum.Text, 6) & "' AND RIGHT(IDSoumission, 2) <> '99' ORDER BY CInt(OrdreSection), NuméroLigne", g_connData, adOpenForwardOnly, adLockReadOnly)

865     Do While Not rstPieces.EOF
870       Call rstPiecesCumulatif.AddNew

875       rstPiecesCumulatif.Fields("IDSoumission") = sNoCumulatif
880       rstPiecesCumulatif.Fields("IDSection") = rstPieces.Fields("IDSection")
885       rstPiecesCumulatif.Fields("NumItem") = rstPieces.Fields("NumItem")
890       rstPiecesCumulatif.Fields("Qté") = rstPieces.Fields("Qté")
895       rstPiecesCumulatif.Fields("Desc_FR") = rstPieces.Fields("Desc_FR")
900       rstPiecesCumulatif.Fields("Desc_EN") = rstPieces.Fields("Desc_EN")
905       rstPiecesCumulatif.Fields("Manufact") = rstPieces.Fields("Manufact")
910       rstPiecesCumulatif.Fields("Prix_list") = rstPieces.Fields("Prix_list")
915       rstPiecesCumulatif.Fields("Escompte") = rstPieces.Fields("Escompte")
920       rstPiecesCumulatif.Fields("Prix_net") = rstPieces.Fields("Prix_net")
925       rstPiecesCumulatif.Fields("IDFRS") = rstPieces.Fields("IDFRS")
930       rstPiecesCumulatif.Fields("Temps") = rstPieces.Fields("Temps")
935       rstPiecesCumulatif.Fields("Temps_Total") = rstPieces.Fields("Temps_Total")
940       rstPiecesCumulatif.Fields("Prix_total") = rstPieces.Fields("Prix_total")
945       rstPiecesCumulatif.Fields("Profit_Argent") = rstPieces.Fields("Profit_Argent")
950       rstPiecesCumulatif.Fields("SousSection") = rstPieces.Fields("SousSection")
955       rstPiecesCumulatif.Fields("OrdreSection") = rstPieces.Fields("OrdreSection")
960       rstPiecesCumulatif.Fields("NuméroLigne") = rstPieces.Fields("NuméroLigne")
965       rstPiecesCumulatif.Fields("PrixOrigine") = rstPieces.Fields("PrixOrigine")
970       rstPiecesCumulatif.Fields("Type") = rstPieces.Fields("Type")
975       rstPiecesCumulatif.Fields("Visible") = rstPieces.Fields("Visible")
980       rstPiecesCumulatif.Fields("Commentaire") = rstPieces.Fields("Commentaire")
985       rstPiecesCumulatif.Fields("Devise") = rstPieces.Fields("Devise")
990       rstPiecesCumulatif.Fields("Provenance") = Right(rstPieces.Fields("IDSoumission"), 2)

995       Call rstPiecesCumulatif.Update

1000      Call rstPieces.MoveNext
1005    Loop

1010    Call rstPiecesCumulatif.Close
1015    Call rstPieces.Close

1020    Set rstSoum = New ADODB.Recordset
1025    Set rstPieces = New ADODB.Recordset
1030    Set rstSoumCumulatif = New ADODB.Recordset
1035    Set rstPiecesCumulatif = New ADODB.Recordset

1040    If bSupprimer = False Then
1045      Call CalculerTotalRecordset(sNoCumulatif)
1050    End If

1055    Exit Sub

AfficherErreur:

1060    woups "FrmProjSoumElec", "RecreerSoumissionCumulatif", Err, Erl
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

