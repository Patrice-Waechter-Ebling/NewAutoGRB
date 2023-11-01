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
   MDIChild        =   -1  'True
   ScaleHeight     =   8265
   ScaleWidth      =   13380
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbOuvertFerme 
      Height          =   315
      ItemData        =   "FrmProjSoumElec.frx":2CFA
      Left            =   4560
      List            =   "FrmProjSoumElec.frx":2D04
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
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   0
      Appearance      =   1
      StartOfWeek     =   36044801
      CurrentDate     =   37761
   End
   Begin MSComCtl2.MonthView mvwDate 
      Height          =   2370
      Left            =   9360
      TabIndex        =   41
      Top             =   720
      Visible         =   0   'False
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   0
      Appearance      =   1
      StartOfWeek     =   36044801
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
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   0
         Appearance      =   1
         StartOfWeek     =   36044801
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
         Picture         =   "FrmProjSoumElec.frx":2D19
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
         Picture         =   "FrmProjSoumElec.frx":58CFB
         Stretch         =   -1  'True
         Top             =   600
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Image imgEU 
         Height          =   1065
         Left            =   6840
         Picture         =   "FrmProjSoumElec.frx":5B18A
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
         ItemData        =   "FrmProjSoumElec.frx":A7EFC
         Left            =   960
         List            =   "FrmProjSoumElec.frx":A7F06
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
      ItemData        =   "FrmProjSoumElec.frx":A7F1E
      Left            =   6480
      List            =   "FrmProjSoumElec.frx":A7F31
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
      ItemData        =   "FrmProjSoumElec.frx":A7F88
      Left            =   4800
      List            =   "FrmProjSoumElec.frx":A7F8A
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
      ItemData        =   "FrmProjSoumElec.frx":A7F8C
      Left            =   3240
      List            =   "FrmProjSoumElec.frx":A7F96
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
      ItemData        =   "FrmProjSoumElec.frx":A7FAE
      Left            =   1080
      List            =   "FrmProjSoumElec.frx":A7FB5
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
' Liste des valeurs du .Text et du .Tag pour chacune des colonnes de lvwSoumission '
'**************************************************************************************'
' COLONNE | TEXTE | TAG '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_QUANTITE | Quantité | ID Section '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_PIECE | Pièce | Sous-section '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_DESCR | Description FR ou EN | Description FR ou EN '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_MANUFACT | Manufacturier | Ordre section '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_PRIX_LIST | Prix listé | Prix d'origine '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_ESCOMPTE | Escompte | '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_PRIX_NET | Prix net | Date de réception '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_DISTRIB | Fournisseur | ID Fournisseur '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_TEMPS | Temps | '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_MONTAGE | Montage | '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_TOTAL | Total | Devise monétaire '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_PROFIT | Profit | EXTRA, RETOUR ou ANNULATION '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_COMMENTAIRE | Commentaire | '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_ID | ID | '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_FACTURATION | Facturation | '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_DATE_COMMANDE | Date Commande | Numéro Retour '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_DATE_REQUISE | Date Requise | Date Retour '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_NOM_COMMANDE | Personne qui a commandé | '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_NO_SEQUENTIEL | Numéro séquentiel du BC | '
'---------------------------|----------------------------|-----------------------------'
'I_COL_SOUM_PROVENANCE | Provenance | '
'---------------------------|----------------------------|-----------------------------'

'Index des colonnes de lvwSoumission
Private Const I_COL_SOUM_QUANTITE As Integer = 0
Private Const I_COL_SOUM_PIECE As Integer = 1
Private Const I_COL_SOUM_DESCR As Integer = 2
Private Const I_COL_SOUM_MANUFACT As Integer = 3
Private Const I_COL_SOUM_PRIX_LIST As Integer = 4
Private Const I_COL_SOUM_ESCOMPTE As Integer = 5
Private Const I_COL_SOUM_PRIX_NET As Integer = 6
Private Const I_COL_SOUM_DISTRIB As Integer = 7
Private Const I_COL_SOUM_TEMPS As Integer = 8
Private Const I_COL_SOUM_MONTAGE As Integer = 9
Private Const I_COL_SOUM_TOTAL As Integer = 10
Private Const I_COL_SOUM_PROFIT As Integer = 11
Private Const I_COL_SOUM_COMMENTAIRE As Integer = 12
Private Const I_COL_SOUM_ID As Integer = 13
Private Const I_COL_SOUM_FACTURATION As Integer = 14
Private Const I_COL_SOUM_DATE_COMMANDE As Integer = 15
Private Const I_COL_SOUM_DATE_REQUISE As Integer = 16
Private Const I_COL_SOUM_NOM_COMMANDE As Integer = 17
Private Const I_COL_SOUM_NO_SEQUENTIEL As Integer = 18
Private Const I_COL_SOUM_PROVENANCE As Integer = 19

Private Const I_COL_SOUMISSION_PROV As Integer = 13

'Index des colonnes de lvwSoumission si les colonnes contenant
'des prix ne sont pas là. (SP est pour Sans Prix)
Private Const I_COL_SOUM_SP_QUANTITE As Integer = 0
Private Const I_COL_SOUM_SP_PIECE As Integer = 1
Private Const I_COL_SOUM_SP_DESCR As Integer = 2
Private Const I_COL_SOUM_SP_MANUFACT As Integer = 3
Private Const I_COL_SOUM_SP_DISTRIB As Integer = 4
Private Const I_COL_SOUM_SP_TEMPS As Integer = 5
Private Const I_COL_SOUM_SP_MONTAGE As Integer = 6
Private Const I_COL_SOUM_SP_COMMENTAIRE As Integer = 7
Private Const I_COL_SOUM_SP_ID As Integer = 8
Private Const I_COL_SOUM_SP_DATE_COMMANDE As Integer = 9
Private Const I_COL_SOUM_SP_DATE_REQUISE As Integer = 10
Private Const I_COL_SOUM_SP_NOM_COMMANDE As Integer = 11
Private Const I_COL_SOUM_SP_NO_SEQUENTIEL As Integer = 12
Private Const I_COL_SOUM_SP_PROVENANCE As Integer = 13

Private Const I_COL_SOUMISSION_SP_PROV As Integer = 8

'Index des colonnes de lvwPieces
Private Const I_COL_PIECES_PIECE_GRB As Integer = 0
Private Const I_COL_PIECES_NO_ITEM As Integer = 1
Private Const I_COL_PIECES_MANUFACT As Integer = 2
Private Const I_COL_PIECES_DESCR_FR As Integer = 3
Private Const I_COL_PIECES_DESCR_EN As Integer = 4

'Index des colonnes de lvwPieceTrouve
Private Const I_COL_RECH_PIECE_GRB As Integer = 0
Private Const I_COL_RECH_NO_ITEM As Integer = 1
Private Const I_COL_RECH_CATEGORIE As Integer = 2
Private Const I_COL_RECH_MANUFACT As Integer = 3
Private Const I_COL_RECH_DESCR_FR As Integer = 4
Private Const I_COL_RECH_DESCR_EN As Integer = 5

'Index des colonnes de lvwFournisseur
Private Const I_COL_FRS_FRS As Integer = 0
Private Const I_COL_FRS_PERS_RESS As Integer = 1
Private Const I_COL_FRS_DATE As Integer = 2
Private Const I_COL_FRS_ENTRER_PAR As Integer = 3
Private Const I_COL_FRS_VALIDE As Integer = 4
Private Const I_COL_FRS_PRIX_LIST As Integer = 5
Private Const I_COL_FRS_ESCOMPTE As Integer = 6
Private Const I_COL_FRS_PRIX_NET As Integer = 7
Private Const I_COL_FRS_PRIX_SP As Integer = 8
Private Const I_COL_FRS_QUOTER As Integer = 9
Private Const I_COL_FRS_STOCK As Integer = 10

'Index des colonnes de lvwModification
Private Const I_COL_MODIF_EMPLOYE As Integer = 0
Private Const I_COL_MODIF_DATE As Integer = 1
Private Const I_COL_MODIF_HEURE As Integer = 2
Private Const I_COL_MODIF_MONTANT As Integer = 3

'Index des colonnes de lvwBavard
Private Const I_COL_SUPP_EMPLOYE As Integer = 0
Private Const I_COL_SUPP_DATE As Integer = 1
Private Const I_COL_SUPP_HEURE As Integer = 2
Private Const I_COL_SUPP_QTE As Integer = 3
Private Const I_COL_SUPP_NO_ITEM As Integer = 4

'Index des transports
Private Const I_TRANS_FAB_GRANBY As Integer = 0
Private Const I_TRANS_CLIENT As Integer = 1

'Index de m_collFloorStock
Private Const I_IDX_FS_DIX_MOINS As Integer = 1
Private Const I_IDX_FS_DIX As Integer = 2
Private Const I_IDX_FS_QUINZE As Integer = 3
Private Const I_IDX_FS_VINGT As Integer = 4
Private Const I_IDX_FS_VINGT_CINQ As Integer = 5
Private Const I_IDX_FS_CINQUANTE As Integer = 6
Private Const I_IDX_FS_CENT As Integer = 7

'Index de cmbChoix
Private Const I_IDX_SOUMISSION As Integer = 0
Private Const I_IDX_PROJET As Integer = 1

'Index de cmbOuvertFerme
Private Const I_CMB_OUVERT As Integer = 0
Private Const I_CMB_TOUS As Integer = 1

'Index de cmbTri
Private Const I_CMB_PIECE_GRB As Integer = 0
Private Const I_CMB_PIECE As Integer = 1
Private Const I_CMB_FABRICANT As Integer = 2
Private Const I_CMB_DESCR_FR As Integer = 3
Private Const I_CMB_DESCR_EN As Integer = 4

'Constante s'il n'y a pas de sous-sections
Private Const S_PAS_SOUS_SECTION As String = "PAS DE SOUS-SECTION"

'Valeur servant au resize du lvwSoumission si le form est agrandi
Private Const I_TOP_AFFICHAGE As Integer = 3000
Private Const I_HEIGHT_AFFICHAGE As Integer = 3930

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
 bChecked As Boolean
 sQuantite As String
 sPiece As String
 sDescr As String
 sManufact As String
 sPrixList As String
 sEscompte As String
 sPrixNet As String
 sFRS As String
 sTemps As String
 sMontage As String
 sTotal As String
 sProfit As String
 sDescrTag As String
 sPrixListTag As String
 sFRSTag As String
 lColor As Long
End Type

'Variables pour la configurations
Private m_sProfit As String
Private m_sCommission As String
Private m_sImprevue As String

'Pour la recherche de pièce dans lvwPieces
Private m_sTri As String

'Pour savoir quelle colonne trier
Private m_iCol As Integer

'Pour savoir si le form a déjà été sur l'événement resize
Private m_bResize As Boolean

'Modes du form
Private m_bModeAjout As Boolean
Private m_bModeAffichage As Boolean

'Pour avoir une sous-section par défaut
Private m_sSousSection As String

'Pour savoir si le form affiche les projets ou les soumissions
Private m_eType As enumType

'Pour savoir si les prix sont cachés ou non
Public m_bDroitPrix As Boolean

'Pour ne pas être obligé d'ouvrir le recordset à chaque fois
Private m_bModifProj As Boolean
Private m_bModifSoum As Boolean
Private m_bModifBonCommande As Boolean

'Variable pour savoir si l'utilisateur a le droit de voir le combo ou non
Private m_bComboChoix As Boolean

Private m_eMode As enumMode

Private m_eLangage As enumLangage

'Pour faire afficher le dernier enregitrement visionné après un ajout ou une
'modification
Private m_sAncienProjSoum As String

Private m_bSupprimer As Boolean

'Pour savoir si il faut calculer le temps mécanique ou non
Public m_bSansTemps As Boolean

'Pour savoir si le lvwFournisseur est affiché après marchandise non utilisée
Private m_bPieceInutile As Boolean

Public m_bAnnulerChemin As Boolean
Public m_sChemin As String

Private m_bRecherchePiece As Boolean

'Pour savoir si le changement de prix a été appelé à partir du bouton "Mauvais Prix"
Private m_bMauvaisPrix As Boolean
Private m_bEnregistrement As Boolean
Private m_collDateSupp As Collection
Private m_collHeureSupp As Collection
Private m_collQteSupp As Collection
Private m_collNoItemSupp As Collection
Private m_bChangementFRS As Boolean

Public m_sTempsDessin As String
Public m_sTempsFabrication As String
Public m_sTempsAssemblage As String
Public m_sTempsProgInterface As String
Public m_sTempsProgAutomate As String
Public m_sTempsProgRobot As String
Public m_sTempsVision As String
Public m_sTempsTest As String
Public m_sTempsInstallation As String
Public m_sTempsMiseService As String
Public m_sTempsFormation As String
Public m_sTempsGestion As String
Public m_sTempsShipping As String

Public m_sNbrePersonne As String
Public m_sTempsHebergement As String
Public m_sTempsRepas As String
Public m_sTempsTransport As String
Public m_sTempsUniteMobile As String
Public m_sPrixEmballage As String

Public m_sTauxHebergement As String
Public m_sTauxRepas As String
Public m_sTauxTransport As String
Public m_sTauxUniteMobile As String

Public m_sTauxDessin As String
Public m_sTauxFabrication As String
Public m_sTauxAssemblage As String
Public m_sTauxProgInterface As String
Public m_sTauxProgAutomate As String
Public m_sTauxProgRobot As String
Public m_sTauxVision As String
Public m_sTauxTest As String
Public m_sTauxInstallation As String
Public m_sTauxMiseService As String
Public m_sTauxFormation As String
Public m_sTauxGestion As String
Public m_sTauxShipping As String

Public m_bTempsDejaOuvert As Boolean

Private m_sTexteRecherche As String
Private m_arr_tyCopie() As tyCopiePiece
Private m_iNbreCopie As Integer
Public m_bModifFournisseurBC As Boolean
Private m_sLiaison As String
Private m_bExtra As Boolean
Private m_bMonthViewHasFocus As Boolean
Public m_bTransfertJobCancel As Boolean
Private m_bChangementChoix As Boolean 'Pour empêcher l'événement cmbOuvertFerme_Click quand cmbChoix change
Public m_bValide As Boolean 'Résultat de frmValiderSuppression


Public intdummie As Integer 'va servir à annuler l'impression des rapports dans la sub ImprimerProjSoum
Public bTrigger As Boolean

'**********************************************************************************************************
'AJOUT PAR GAÉTAN GINGRAS 0  FÉVRIER 2010
'**********************************************************************************************************
Public bFlag As Boolean 'pour garder en mémoire si on désir achiffer les dates de réception et de commande
'**********************************************************************************************************
 
Public Function PeutFermer() As Boolean
 
 On Error GoTo Oups
 
 If m_eMode = MODE_INACTIF Then
 PeutFermer = True
 Else
 PeutFermer = False
 End If

 Exit Function

Oups:

 wOups "frmProjSoumElec", "PeutFermer", Err, Err.number, Err.Description
End Function

Private Sub InitialiserVariables(ByVal sNoProjSoum As String)

 On Error GoTo Oups

 'Initialisation des variables comprises dans la configuration
 Dim rstConfig As ADODB.Recordset
 Dim rstProjSoum As ADODB.Recordset

 Set rstProjSoum = New ADODB.Recordset

 If m_eType = TYPE_PROJET Then
 Call rstProjSoum.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstProjSoum.Open("SELECT * FROM GrbSoumissionElec WHERE IDSoumission = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)
 End If

 If Not rstProjSoum.EOF Then
 m_sProfit = rstProjSoum.Fields("Profit")
  m_sCommission = rstProjSoum.Fields("Commission")
  m_sImprevue = rstProjSoum.Fields("Imprevue")
  Else
  Set rstConfig = New ADODB.Recordset

  Call rstConfig.Open("SELECT * FROM GrbConfig", g_connData, adOpenDynamic, adLockOptimistic)

  m_sProfit = rstConfig.Fields("ProfitElec")
  m_sCommission = rstConfig.Fields("Commission")
  m_sImprevue = rstConfig.Fields("Imprévus")

Call rstConfig.Close
1 Set rstConfig = Nothing
End If

Call rstProjSoum.Close
Set rstProjSoum = Nothing

Exit Sub

Oups:

wOups "frmProjSoumElec", "InitialiserVariables", Err, Err.number, Err.Description
End Sub

Private Sub ActiverBoutonsGroupe()

 On Error GoTo Oups

 'Activation des boutons d'après le groupe
 Dim bModif As Boolean

 'Si l'utilisateur a le droit d'affichage sur les projets et les soumissions
 If g_bAffichageProjetsElec = True And g_bAffichageSoumissionsElec = True Then
 'On affiche cmbChoix
 cmbChoix.Visible = True

 'Cette variable sert à savoir si l'utilisateur a le droit de voir le combo
 m_bComboChoix = True

 'Type d'affichage
 m_eType = TYPE_PROJET

 'Champs pour la modification
 bModif = g_bModificationProjetsElec
 Else
 'On cache cmbChoix
 cmbChoix.Visible = False

 'Cette variable sert à savoir si l'utilisateur a le droit de voir le combo
 m_bComboChoix = False

 'Si l'utilisateur a le droit d'affichage sur les projets
 If g_bAffichageProjetsElec = True Then
 'Le seul choix possible est Projet
  txtChoix.Text = "Projet"

 'Le type d'affichage
  m_eType = TYPE_PROJET

 'Champs pour la modification
  bModif = g_bModificationProjetsElec
  Else
 'Le seul choix possible est Soumission
  txtChoix.Text = "Soumission"

 'Type d'affichage
  m_eType = TYPE_SOUMISSION

 'Champs pour la modification
  bModif = g_bModificationSoumissionsElec
  End If
10 End If

m_bModifProj = g_bModificationProjetsElec
m_bModifSoum = g_bModificationSoumissionsElec
m_bModifBonCommande = g_bModificationBC
m_bSupprimer = g_bSuppressionProjets

Cmdajouter.Enabled = bModif
cmdsupprimer.Enabled = bModif
cmdModifier.Enabled = bModif
cmdCopier.Enabled = bModif
cmdCreerProjet.Enabled = bModif
cmdBonCommande.Enabled = m_bModifBonCommande
cmdImprimer.Enabled = bModif
1  cmdDemande.Enabled = bModif
cmdAnglaisFrancais.Enabled = bModif
 cmdExtra.Enabled = bModif
cmdSupprimerFRS.Visible = g_bModificationCatalogueElec
 cmdRetour.Enabled = g_bModificationRetourMarchandise
cmdReception.Enabled = g_bModificationReception

 Exit Sub

Oups:

1  wOups "frmProjSoumElec", "ActiverBoutonsGroupe", Err, Err.number, Err.Description
End Sub

Private Sub AfficherProjSoum(ByVal sNoProjSoum As String)

 On Error GoTo Oups

 m_bPieceInutile = False
 m_bChangementFRS = False
 m_bRecherchePiece = False

 'Remet en mode affichage le projet ou la soumission voulue
 m_bModeAffichage = True
 
 'Vide les champs
 Call ViderChamps
 
 'Rempli le combo
 Call RemplirComboProjSoum(sNoProjSoum)
 
 'Barre les champs
 Call BarrerChamps(True)
 
 lvwSoumission.Top = I_TOP_AFFICHAGE
 lvwSoumission.Height = Me.Height - I_HEIGHT_AFFICHAGE

 Exit Sub

Oups:

  wOups "frmProjSoumElec", "AfficherProjSoum", Err, Err.number, Err.Description
End Sub

Private Sub AfficherControles(ByVal eMode As enumMode)

 On Error GoTo Oups

 'Affichage des boutons selon si c'est un ajout/modif ou un affichage
 Dim bAjouter As Boolean
 Dim bModifier As Boolean
 Dim bSupprimer As Boolean
 Dim bEnregistrer As Boolean
 Dim bAnnuler As Boolean
 Dim bFermer As Boolean
 Dim bImprimer As Boolean
 Dim bCmbClient As Boolean
 Dim bCmbContact As Boolean
 Dim bCmbProjSoum As Boolean
  Dim bCmbTransport As Boolean
  Dim bCmbChoix As Boolean
  Dim bCmbOuvertFerme As Boolean
  Dim bSection As Boolean
  Dim bPieces As Boolean
  Dim bDate As Boolean
  Dim bTexte As Boolean
  Dim bCreerProjet As Boolean
10 Dim bHistorique As Boolean
Dim bCopier As Boolean
Dim bBonCommande As Boolean
Dim bTri As Boolean
Dim bDemande As Boolean
Dim bExtra As Boolean
Dim bCatalogue As Boolean
Dim bBrowseChemin As Boolean
Dim bInutile As Boolean
Dim bMauvaisPrix As Boolean
Dim bRapportFact As Boolean
Dim bDateFacture As Boolean
1  Dim bSortiMagasin As Boolean
Dim bRetour As Boolean
 Dim bForfait As Boolean
Dim bExporter As Boolean
 Dim bReception As Boolean
Dim bAnglaisFrancais As Boolean
 Dim bRechercheClient As Boolean
 
1  m_eMode = eMode
 
 Select Case eMode
 Case MODE_AJOUT_MODIF:
 bEnregistrer = True
 bAnnuler = True

 bSection = True
 bPieces = True
 bTexte = True
 bTri = True

 If (m_eType = TYPE_SOUMISSION) Or (m_eType = TYPE_PROJET And Mid$(txtNoProjSoum.Text, 3, 1) <> "3") Then
 bCmbClient = True
 bCmbContact = True
 bRechercheClient = True
 End If

 bCmbTransport = True
 bDate = True
 bCatalogue = True
 bBrowseChemin = True
 bMauvaisPrix = True
 bForfait = True

 If m_eType = TYPE_PROJET Then
 bInutile = True

 If g_bModificationReception = True Then
 bSortiMagasin = True
 End If

 If g_bModificationFacturation = True Then
 bDateFacture = True
 End If
 End If
 
 Case MODE_INACTIF:
 bModifier = True
 bFermer = True
 bImprimer = True
 bCmbProjSoum = True
 bCmbChoix = True
 bCmbOuvertFerme = True
 bHistorique = True
 bDemande = True
 bExporter = True
 bAnglaisFrancais = True
 bAjouter = True
 
 If m_eType = TYPE_PROJET Then
 bBonCommande = True
4 bExtra = True

4 If g_bModificationRetourMarchandise = True Then
4 bRetour = True
4 End If

4 If g_bModificationFacturation = True Then
4 bRapportFact = True
4 End If

4 If g_bModificationReception = True Then
4 bReception = True
4 End If
 
4 If m_bSupprimer = True Then
4  bSupprimer = True
4  End If
4  Else
4  bSupprimer = True
4  bCopier = True
 
4  If VerifierSiDejaProjet = False Then
4  bCreerProjet = True
4  End If
50 End If
50 End Select
 
 Cmdajouter.Visible = bAjouter
 cmdModifier.Visible = bModifier
 cmdsupprimer.Visible = bSupprimer
 cmdEnregistrer.Visible = bEnregistrer
 cmdAnnuler.Visible = bAnnuler
 Cmdfermer.Visible = bFermer
 cmdImprimer.Visible = bImprimer
 cmdRapportFACT.Visible = bRapportFact
 cmdDate.Visible = bDate
 cmdTexte.Visible = bTexte
5  cmdHistorique.Visible = bHistorique
5  cmdCopier.Visible = bCopier
5  cmdBonCommande.Visible = bBonCommande
5  cmdCreerProjet.Visible = bCreerProjet
5  cmdDemande.Visible = bDemande
5  cmdExtra.Visible = bExtra
5  cmdCatalogue.Visible = bCatalogue
5  cmdBrowse.Visible = bBrowseChemin
60 cmdMaterielInutile.Visible = bInutile
60 cmdMauvaisPrix.Visible = bMauvaisPrix
  cmdSortieMagasin.Visible = bSortiMagasin
  cmdRetour.Visible = bRetour
  cmdForfait.Visible = bForfait
  cmdEffacerForfait.Visible = bForfait
  cmdExport.Visible = bExporter
  cmdReception.Visible = bReception
  cmdAnglaisFrancais.Visible = bAnglaisFrancais

  lblDateFacturation.Visible = bDateFacture
  txtDateFacturation.Visible = bDateFacture
  cmdDateFacturation.Visible = bDateFacture
 
6  cmbclient.Visible = bCmbClient
6  txtClient.Visible = Not bCmbClient

6  cmbContact.Visible = bCmbContact
6  txtcontact.Visible = Not bCmbContact
 
6  cmbtransport.Visible = bCmbTransport
6  txtTransport.Visible = Not bCmbTransport
 
 'Si on a le droit d'afficher le combo
6  If m_bComboChoix = True Then
6  cmbChoix.Visible = bCmbChoix
70 txtChoix.Visible = Not bCmbChoix
70 End If

  cmbOuvertFerme.Visible = bCmbOuvertFerme
 
  cmbProjSoum.Visible = bCmbProjSoum
  txtNoProjSoum.Visible = Not bCmbProjSoum
 
  lblSections.Visible = bSection
  cmbSections.Visible = bSection
  cmdAjouterSection.Visible = bSection

  lblPiece.Visible = bPieces
  cmbPieces.Visible = bPieces
  lvwPieces.Visible = bPieces

  lblTri.Visible = bTri
   cmbTri.Visible = bTri
   cmdTri.Visible = bTri
7  cmdRafraichir.Visible = bTri
 
7  fraPrix.Visible = m_bDroitPrix

7  cmdRechercherClient.Visible = bRechercheClient

7  Exit Sub

Oups:

7  wOups "frmProjSoumElec", "AfficherControles", Err, Err.number, Err.Description
End Sub

Private Sub cmbChoix_Click()

 On Error GoTo Oups

 Dim bModif As Boolean
 Dim iCmbOuvertFerme As Integer
 
 Screen.MousePointer = vbHourglass
 
 txtChoix.Text = cmbChoix.Text

 'Mets les CheckBoxes sur le ListView
 lvwSoumission.CheckBoxes = True

 If cmbChoix.ListIndex = I_IDX_SOUMISSION Then
 'Change le type
 m_eType = TYPE_SOUMISSION
 
 m_bChangementChoix = True
 
 iCmbOuvertFerme = cmbOuvertFerme.ListIndex
 
 Call cmbOuvertFerme.Clear

  Call cmbOuvertFerme.AddItem("Ouvertes")
  Call cmbOuvertFerme.AddItem("Toutes")

  cmbOuvertFerme.ListIndex = iCmbOuvertFerme
 
  m_bChangementChoix = False
 
  bModif = m_bModifSoum
 
 'Cache la soumission
  lblNoSoumission.Visible = False
  txtNoSoumission.Visible = False

 'Cache Prix Réception
  lblPrixReception.Visible = False
txtPrixReception.Visible = False

 'Cache Prix Soumission
1 lblPrixSoumission.Visible = False
 txtPrixSoumission.Visible = False
Else
 'Change le type
 m_eType = TYPE_PROJET
 
 m_bChangementChoix = True
 
 iCmbOuvertFerme = cmbOuvertFerme.ListIndex

 Call cmbOuvertFerme.Clear

 Call cmbOuvertFerme.AddItem("Ouverts")
 Call cmbOuvertFerme.AddItem("Tous")

 cmbOuvertFerme.ListIndex = iCmbOuvertFerme
 
 m_bChangementChoix = False
 
bModif = m_bModifProj

 'Affiche la soumission
 lblNoSoumission.Visible = True
 txtNoSoumission.Visible = True

 'Affiche Prix Réception
 lblPrixReception.Visible = True
 txtPrixReception.Visible = True

 'Affiche Prix Soumission
 lblPrixSoumission.Visible = True
 txtPrixSoumission.Visible = True

1  txtDateFacturation.Text = ConvertDate(Date)
 End If
 
 'Active ou désactive les boutons de modification selon
 'le groupe auquel l'utilisateur appartient
 cmdModifier.Enabled = bModif
cmdsupprimer.Enabled = bModif
Cmdajouter.Enabled = bModif
cmdCopier.Enabled = bModif
cmdCreerProjet.Enabled = bModif
cmdBonCommande.Enabled = m_bModifBonCommande
cmdImprimer.Enabled = bModif
cmdAnglaisFrancais.Enabled = bModif
cmdDemande.Enabled = bModif
cmdExtra.Enabled = bModif
 
 'Ajoute les colonnes selon le groupe
Call RemplirColonnes
 
2  m_bModeAffichage = True
 
 'Vide les champs
Call ViderChamps
 
 'Barre les champs
2  Call BarrerChamps(True)
 
 'Rempli le combo
Call RemplirComboProjSoum(vbNullString)

 'Affiche les controles pour le mode inactif
2  Call AfficherControles(MODE_INACTIF)
 
Call PositionnerBoutons

2  Screen.MousePointer = vbDefault

Exit Sub

Oups:

30 wOups "frmProjSoumElec", "cmbChoix_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmbclient_Click()

 On Error GoTo Oups

 'Rempli le combo des contacts selon le client choisi
 Call RemplirComboContacts

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "cmbclient_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmbOuvertFerme_Click()
 
 On Error GoTo Oups

 If cmbChoix.ListIndex <> -1 Then
 If m_bChangementChoix = False Then
 Call RemplirComboProjSoum("")
 End If
 End If

 Exit Sub

Oups:

 wOups "FrmProjSoumElec", "cmbOuvertFerme_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmbPieces_Click()

 On Error GoTo Oups

 'Rempli lvwPieces selon la catégorie de pièce choisie
 Call RemplirListViewPieces

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "cmbPieces_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmbProjSoum_Click()

 On Error GoTo Oups

 Dim rstProjSoum As ADODB.Recordset
 Dim rstOuvert As ADODB.Recordset
 Dim iCompteur As Integer
 Dim sNomClient As String
 Dim sNomContact As String
 Dim sNumero As String
 Dim sTransport As String
 Dim bTrouve As Boolean
 
 Screen.MousePointer = vbHourglass

 m_bRecherchePiece = False
  m_bChangementFRS = False
  m_bPieceInutile = False

  If cmbProjSoum.Text <> "" Then
  sNumero = txtNoProjSoum.Text

  txtNoProjSoum.Text = cmbProjSoum.Text

  Call InitialiserVariables(txtNoProjSoum.Text)

  If m_bEnregistrement = False Then
  m_eLangage = FRANCAIS

 cmdAnglaisFrancais.Caption = "Anglais"
1 End If
 
 Set rstProjSoum = New ADODB.Recordset
 
 If m_eType = TYPE_SOUMISSION Then
 Call rstProjSoum.Open("SELECT * FROM GrbSoumissionElec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstProjSoum.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
 End If

 If rstProjSoum.Fields("Modification") = True And rstProjSoum.Fields("Par") = g_sEmploye Then
 cmdReset.Visible = True
 End If

 Call InitialiserTempsTaux(False)

If m_eType = TYPE_SOUMISSION Then
 'Si la soumission n'est pas assigné à un projet
 If VerifierSiDejaProjet = False Then
 'On affiche le bouton cmdCreerProjet
 cmdCreerProjet.Visible = True
 Else
 cmdCreerProjet.Visible = False
 End If
 End If
 
 'Rempli les valeurs de la soumission ou du projet sélectionné
1  Call RemplirProjSoum

 'Le temps calculé dans le projet est le temps réel, c'est pourquoi il faut le recalculer
 'puisque le temps varie souvent
 If m_eType = TYPE_PROJET Then
 Set rstOuvert = New ADODB.Recordset

 Call rstOuvert.Open("SELECT * FROM GrbProjSoum WHERE IDProjSoum = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

 If rstOuvert.Fields("Ouvert") = True Then
 m_bModeAffichage = False

 Call CalculerPrix

 m_bModeAffichage = True

 rstProjSoum.Fields("total_Commission") = txtCommission.Text
 rstProjSoum.Fields("Total_Profit") = txtProfit.Text
 rstProjSoum.Fields("PrixTotal") = txtPrixTotal.Text
 rstProjSoum.Fields("Total_piece") = txtTotalPieces.Text
 rstProjSoum.Fields("total_imprevue") = txtImprevus.Text
 rstProjSoum.Fields("Total_Temps") = txtTotalTemps.Text

 Call rstProjSoum.Update
 End If
 End If
 
Call rstProjSoum.Close

 sNomClient = txtClient.Text
sNomContact = txtcontact.Text
 sTransport = txtTransport.Text
 
 'Pour choisir le bon client dans le combo des clients
For iCompteur = 0 To cmbclient.ListCount - 1
If cmbclient.LIST(iCompteur) = sNomClient Then
 cmbclient.ListIndex = iCompteur

 bTrouve = True
 
 Exit For
 End If
 Next
 
 If bTrouve = False Then
 Call RemplirComboClients(vbNullString)

 For iCompteur = 0 To cmbclient.ListCount - 1
 If cmbclient.LIST(iCompteur) = sNomClient Then
 cmbclient.ListIndex = iCompteur

 Exit For
 End If
 Next
 End If
 
 'Pour choisir le bon contact dans le combo des contacts
For iCompteur = 0 To cmbContact.ListCount - 1
 If cmbContact.LIST(iCompteur) = sNomContact Then
 cmbContact.ListIndex = iCompteur
 
 Exit For
 End If
4 Next
 
 'Pour choisir le bon transport dans le combo des transports
4 For iCompteur = 0 To cmbtransport.ListCount - 1
4 If cmbtransport.LIST(iCompteur) = sTransport Then
4 cmbtransport.ListIndex = iCompteur
 
4 Exit For
4 End If
4 Next
4 End If

4 Call CalculerPrixReception

4 If m_eType = TYPE_PROJET Then
4 rstProjSoum.CursorLocation = adUseServer

4  Call rstProjSoum.Open("SELECT PrixRéception FROM GrbProjetElec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
4  rstProjSoum.Fields("PrixRéception") = txtPrixReception.Text

4  Call rstProjSoum.Update

4  Call rstProjSoum.Close
4  End If

4  If m_bSansTemps = True Then
4  tmrTemps.Enabled = True
4  Else
50 tmrTemps.Enabled = False
5 lblPasTemps.Visible = False
 End If
 
 Set rstProjSoum = Nothing
 
 Screen.MousePointer = vbDefault

 Exit Sub

Oups:

 woups"frmProjSoumElec", "cmbProjSoum_Click", Err, Erl, txtNoProjSoum.Text)
End Sub

Private Sub InitialiserTempsTaux(ByVal bEmpty As Boolean)

 On Error GoTo Oups

 'Pour initialiser les temps et les taux horaires
 Dim rstProjSoum As ADODB.Recordset
 Dim sTable As String
 Dim sChamps As String

 m_bTempsDejaOuvert = False

 If bEmpty = True Then
 m_sTempsDessin = "0"
 m_sTempsFabrication = "0"
 m_sTempsAssemblage = "0"
 m_sTempsProgInterface = "0"
 m_sTempsProgAutomate = "0"
  m_sTempsProgRobot = "0"
  m_sTempsVision = "0"
  m_sTempsTest = "0"
  m_sTempsInstallation = "0"
  m_sTempsMiseService = "0"
  m_sTempsFormation = "0"
  m_sTempsGestion = "0"
  m_sTempsShipping = "0"

m_sNbrePersonne = "0"
1 m_sTempsHebergement = "0"
 m_sTempsRepas = "0"
 m_sTempsTransport = "0"
 m_sTempsUniteMobile = "0"
 m_sPrixEmballage = "0"
 m_sTauxHebergement1 = "0"
 m_sTauxHebergement2 = "0"
 m_sTauxRepas = "0"
 m_sTauxTransport = "0"
 m_sTauxUniteMobile = "0"

 m_sTauxDessin = "0"
m_sTauxFabrication = "0"
 m_sTauxAssemblage = "0"
 m_sTauxProgInterface = "0"
 m_sTauxProgAutomate = "0"
 m_sTauxProgRobot = "0"
 m_sTauxVision = "0"
 m_sTauxTest = "0"
1  m_sTauxInstallation = "0"
 m_sTauxMiseService = "0"
 m_sTauxFormation = "0"
 m_sTauxGestion = "0"
 m_sTauxShipping = "0"
Else
 If m_eType = TYPE_PROJET Then
 sTable = "GrbProjetElec"
 sChamps = "IDProjet"
 Else
 sTable = "GrbSoumissionElec"
 sChamps = "IDSoumission"
 End If

Set rstProjSoum = New ADODB.Recordset

 Call rstProjSoum.Open("SELECT * FROM " & sTable & " WHERE " & sChamps & " = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

If m_eType = TYPE_SOUMISSION Then
 If Not IsNull(rstProjSoum.Fields("TempsDessin")) Then
 m_sTempsDessin = rstProjSoum.Fields("TempsDessin")
 Else
 m_sTempsDessin = "0"
 End If

 If Not IsNull(rstProjSoum.Fields("TempsFabrication")) Then
 m_sTempsFabrication = rstProjSoum.Fields("TempsFabrication")
 Else
 m_sTempsFabrication = "0"
 End If

 If Not IsNull(rstProjSoum.Fields("TempsAssemblage")) Then
 m_sTempsAssemblage = rstProjSoum.Fields("TempsAssemblage")
 Else
 m_sTempsAssemblage = "0"
 End If

 If Not IsNull(rstProjSoum.Fields("TempsProgInterface")) Then
 m_sTempsProgInterface = rstProjSoum.Fields("TempsProgInterface")
 Else
 m_sTempsProgInterface = "0"
 End If

 If Not IsNull(rstProjSoum.Fields("TempsProgAutomate")) Then
 m_sTempsProgAutomate = rstProjSoum.Fields("TempsProgAutomate")
 Else
 m_sTempsProgAutomate = "0"
 End If

 If Not IsNull(rstProjSoum.Fields("TempsProgRobot")) Then
4 m_sTempsProgRobot = rstProjSoum.Fields("TempsProgRobot")
4 Else
4 m_sTempsProgRobot = "0"
4 End If

4 If Not IsNull(rstProjSoum.Fields("TempsVision")) Then
4 m_sTempsVision = rstProjSoum.Fields("TempsVision")
4 Else
4 m_sTempsVision = "0"
4 End If

4 If Not IsNull(rstProjSoum.Fields("TempsTest")) Then
4 m_sTempsTest = rstProjSoum.Fields("TempsTest")
4  Else
4  m_sTempsTest = "0"
4  End If

4  If Not IsNull(rstProjSoum.Fields("TempsInstallation")) Then
4  m_sTempsInstallation = rstProjSoum.Fields("TempsInstallation")
4  Else
4  m_sTempsInstallation = "0"
4  End If

50 If Not IsNull(rstProjSoum.Fields("TempsMiseService")) Then
 m_sTempsMiseService = rstProjSoum.Fields("TempsMiseService")
 Else
 m_sTempsMiseService = "0"
 End If

 If Not IsNull(rstProjSoum.Fields("TempsFormation")) Then
 m_sTempsFormation = rstProjSoum.Fields("TempsFormation")
 Else
 m_sTempsFormation = "0"
 End If

 If Not IsNull(rstProjSoum.Fields("TempsGestion")) Then
 m_sTempsGestion = rstProjSoum.Fields("TempsGestion")
5  Else
5  m_sTempsGestion = "0"
5  End If

5  If Not IsNull(rstProjSoum.Fields("TempsShipping")) Then
5  m_sTempsShipping = rstProjSoum.Fields("TempsShipping")
5  Else
5  m_sTempsShipping = "0"
5  End If
60 Else
  Call InitialiserTempsReel
  End If

  If Not IsNull(rstProjSoum.Fields("TauxDessin")) Then
  m_sTauxDessin = rstProjSoum.Fields("TauxDessin")
  Else
  m_sTauxDessin = "0"
  End If

  If Not IsNull(rstProjSoum.Fields("TauxFabrication")) Then
  m_sTauxFabrication = rstProjSoum.Fields("TauxFabrication")
  Else
  m_sTauxFabrication = "0"
6  End If

6  If Not IsNull(rstProjSoum.Fields("TauxAssemblage")) Then
6  m_sTauxAssemblage = rstProjSoum.Fields("TauxAssemblage")
6  Else
6  m_sTauxAssemblage = "0"
6  End If

6  If Not IsNull(rstProjSoum.Fields("TauxProgInterface")) Then
6  m_sTauxProgInterface = rstProjSoum.Fields("TauxProgInterface")
70 Else
  m_sTauxProgInterface = "0"
  End If

  If Not IsNull(rstProjSoum.Fields("TauxProgAutomate")) Then
  m_sTauxProgAutomate = rstProjSoum.Fields("TauxProgAutomate")
  Else
  m_sTauxProgAutomate = "0"
  End If

  If Not IsNull(rstProjSoum.Fields("TauxProgRobot")) Then
  m_sTauxProgRobot = rstProjSoum.Fields("TauxProgRobot")
  Else
  m_sTauxProgRobot = "0"
   End If

   If Not IsNull(rstProjSoum.Fields("TauxVision")) Then
7  m_sTauxVision = rstProjSoum.Fields("TauxVision")
7  Else
7  m_sTauxVision = "0"
7  End If

7  If Not IsNull(rstProjSoum.Fields("TauxTest")) Then
7  m_sTauxTest = rstProjSoum.Fields("TauxTest")
80 Else
  m_sTauxTest = "0"
  End If

  If Not IsNull(rstProjSoum.Fields("TauxInstallation")) Then
  m_sTauxInstallation = rstProjSoum.Fields("TauxInstallation")
  Else
  m_sTauxInstallation = "0"
  End If

  If Not IsNull(rstProjSoum.Fields("TauxMiseService")) Then
  m_sTauxMiseService = rstProjSoum.Fields("TauxMiseService")
  Else
  m_sTauxMiseService = "0"
   End If

   If Not IsNull(rstProjSoum.Fields("TauxFormation")) Then
   m_sTauxFormation = rstProjSoum.Fields("TauxFormation")
   Else
8  m_sTauxFormation = "0"
8  End If

8  If Not IsNull(rstProjSoum.Fields("TauxGestion")) Then
8  m_sTauxGestion = rstProjSoum.Fields("TauxGestion")
90 Else
  m_sTauxGestion = "0"
  End If

  If Not IsNull(rstProjSoum.Fields("TauxShipping")) Then
  m_sTauxShipping = rstProjSoum.Fields("TauxShipping")
  Else
  m_sTauxShipping = "0"
  End If

  If m_eType = TYPE_PROJET Then
  m_sNbrePersonne = "0"
  m_sTempsHebergement = "0"
  m_sTempsRepas = "0"
 m_sTempsTransport = "0"
   m_sTempsUniteMobile = "0"
 Else
   If Not IsNull(rstProjSoum.Fields("NbrePersonne")) Then
 m_sNbrePersonne = rstProjSoum.Fields("NbrePersonne")
   Else
 m_sNbrePersonne = "0"
9  End If
 
 If Not IsNull(rstProjSoum.Fields("TempsHebergement")) Then
 m_sTempsHebergement = rstProjSoum.Fields("TempsHebergement")
1 Else
 m_sTempsHebergement = "0"
 End If

1 If Not IsNull(rstProjSoum.Fields("TempsRepas")) Then
 m_sTempsRepas = rstProjSoum.Fields("TempsRepas")
1 Else
 m_sTempsRepas = "0"
1 End If

 If Not IsNull(rstProjSoum.Fields("TempsTransport")) Then
 m_sTempsTransport = rstProjSoum.Fields("TempsTransport")
10  Else
10  m_sTempsTransport = "0"
10  End If

10  If Not IsNull(rstProjSoum.Fields("TempsUniteMobile")) Then
10  m_sTempsUniteMobile = rstProjSoum.Fields("TempsUniteMobile")
10  Else
10  m_sTempsUniteMobile = "0"
10  End If
110 End If

11 If Not IsNull(rstProjSoum.Fields("PrixEmballage")) Then
1 m_sPrixEmballage = rstProjSoum.Fields("PrixEmballage")
1 Else
1 m_sPrixEmballage = "0"
1 End If

1 If m_eType = TYPE_PROJET Then
1 m_sTauxHebergement1 = "0"
1 m_sTauxHebergement2 = "0"
1 m_sTauxRepas = "0"
1 m_sTauxTransport = "0"
1 m_sTauxUniteMobile = "0"
11  Else
1 If Not IsNull(rstProjSoum.Fields("TauxHebergement1")) Then
 m_sTauxHebergement1 = rstProjSoum.Fields("TauxHebergement1")
1 Else
 m_sTauxHebergement1 = "0"
1 End If

 If Not IsNull(rstProjSoum.Fields("TauxHebergement2")) Then
11  m_sTauxHebergement2 = rstProjSoum.Fields("TauxHebergement2")
 Else
 m_sTauxHebergement2 = "0"
1 End If

1 If Not IsNull(rstProjSoum.Fields("TauxRepas")) Then
1 m_sTauxRepas = rstProjSoum.Fields("TauxRepas")
1 Else
1 m_sTauxRepas = "0"
1 End If

1 If Not IsNull(rstProjSoum.Fields("TauxTransport")) Then
1 m_sTauxTransport = rstProjSoum.Fields("TauxTransport")
1 Else
1 m_sTauxTransport = "0"
1 End If

1 If Not IsNull(rstProjSoum.Fields("TauxUniteMobile")) Then
1 m_sTauxUniteMobile = rstProjSoum.Fields("TauxUniteMobile")
1 Else
1 m_sTauxUniteMobile = "0"
1 End If
12  End If

1 Call rstProjSoum.Close
130 Set rstProjSoum = Nothing
130 End If

13 Exit Sub

Oups:

13 wOups "frmProjSoumElec", "InitialiserTempsTaux", Err, Err.number, Err.Description
End Sub

Private Sub InitialiserTempsReel()

 On Error GoTo Oups

 Dim rstPunch As ADODB.Recordset
 Dim sDateDebut As String
 Dim sDateFin As String
 Dim sTotal As String
 Dim sFilterNoProjet As String
 
 If Right$(txtNoProjSoum.Text, 2) = "99" Then
 sFilterNoProjet = "LEFT(NoProjet, 6) = '" & Left$(txtNoProjSoum.Text, 6) & "'"
 Else
 sFilterNoProjet = "NoProjet = '" & txtNoProjSoum.Text & "'"
 End If

  sDateDebut = "TIMESERIAL(Left(GrbPunch.HeureDébut,2),RIGHT(GrbPunch.HeureDébut,2),0)"

  sDateFin = "TIMESERIAL(Left(GrbPunch.HeureFin,2),RIGHT(GrbPunch.HeureFin,2),0)"

  sTotal = "(SUM(" & sDateFin & " - " & sDateDebut & ")* 24) As Total"

  Set rstPunch = New ADODB.Recordset

  Call rstPunch.Open("SELECT Type, " & sTotal & " FROM GrbPunch WHERE " & sFilterNoProjet & " And HeureFin is not null AND HeureDébut is not null GROUP BY Type", g_connData, adOpenDynamic, adLockOptimistic)

  m_sTempsDessin = "0"
  m_sTempsFabrication = "0"
  m_sTempsAssemblage = "0"
10 m_sTempsProgInterface = "0"
m_sTempsProgAutomate = "0"
m_sTempsProgRobot = "0"
m_sTempsVision = "0"
m_sTempsTest = "0"
m_sTempsInstallation = "0"
m_sTempsMiseService = "0"
m_sTempsFormation = "0"
m_sTempsGestion = "0"
m_sTempsShipping = "0"

Do While Not rstPunch.EOF
 If Not IsNull(rstPunch.Fields("Total")) Then
 Select Case rstPunch.Fields("Type")
 Case "Dessin": m_sTempsDessin = Round(rstPunch.Fields("Total"), 2)
 Case "Fabrication": m_sTempsFabrication = Round(rstPunch.Fields("Total"), 2)
 Case "Assemblage": m_sTempsAssemblage = Round(rstPunch.Fields("Total"), 2)
 Case "ProgInterface": m_sTempsProgInterface = Round(rstPunch.Fields("Total"), 2)
 Case "ProgAutomate": m_sTempsProgAutomate = Round(rstPunch.Fields("Total"), 2)
 Case "ProgRobot": m_sTempsProgRobot = Round(rstPunch.Fields("Total"), 2)
 Case "Vision": m_sTempsVision = Round(rstPunch.Fields("Total"), 2)
1  Case "Test": m_sTempsTest = Round(rstPunch.Fields("Total"), 2)
 Case "Installation": m_sTempsInstallation = Round(rstPunch.Fields("Total"), 2)
 Case "MiseService": m_sTempsMiseService = Round(rstPunch.Fields("Total"), 2)
 Case "Formation": m_sTempsFormation = Round(rstPunch.Fields("Total"), 2)
 Case "Gestion": m_sTempsGestion = Round(rstPunch.Fields("Total"), 2)
 Case "Shipping": m_sTempsShipping = Round(rstPunch.Fields("Total"), 2)
 End Select
 End If

 Call rstPunch.MoveNext
Loop

Call rstPunch.Close

Set rstPunch = Nothing

Exit Sub

Oups:

2  wOups "frmProjSoumElecTemps", "AfficherTempsReels", Err, Err.number, Err.Description
End Sub

Private Sub cmbProjSoum_KeyUp(KeyCode As Integer, Shift As Integer)
 
 On Error GoTo Oups

 Dim iCompteur As Integer
 
 For iCompteur = 0 To cmbProjSoum.ListCount - 1
 If UCase(cmbProjSoum.LIST(iCompteur)) = UCase(cmbProjSoum.Text) Then
 cmbProjSoum.ListIndex = iCompteur
 
 Exit For
 End If
 Next

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "cmbProjSoum_KeyUp", Err, Err.number, Err.Description
End Sub

Private Sub cmdAjouterSection_Click()

 On Error GoTo Oups

 'Affiche le form frmSoumissionSection
 Call OuvrirForm(frmSoumissionSectionElec, True)

 'Après que l'utilisateur a refermé le form, on rafraichi le
 'contenu du combo
 Call RemplirComboSections

 Call UpdateOrdre

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "cmdAjouterSection_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnglaisFrancais_Click()

 On Error GoTo Oups

 Screen.MousePointer = vbHourglass

 If cmdAnglaisFrancais.Caption = "Anglais" Then
 m_eLangage = ANGLAIS
 
 cmdAnglaisFrancais.Caption = "Français"
 Else
 m_eLangage = FRANCAIS
 
 cmdAnglaisFrancais.Caption = "Anglais"
 End If

 Call UpdateDescription
 
 Call RemplirComboSections
 
  Call UpdateOrdre
 
  Screen.MousePointer = vbDefault

  Exit Sub

Oups:

  wOups "frmProjSoumElec", "cmdAnglaisFrancais_Click", Err, Err.number, Err.Description
End Sub

Private Sub UpdateDescription()

 On Error GoTo Oups

 Dim rstProjSoum As ADODB.Recordset
 Dim rstPieceElec As ADODB.Recordset

 Set rstProjSoum = New ADODB.Recordset
 Set rstPieceElec = New ADODB.Recordset

 If m_eType = TYPE_PROJET Then
 Call rstProjSoum.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstProjSoum.Open("SELECT * FROM GrbSoumission_Pieces WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
 End If

 Do While Not rstProjSoum.EOF
  Call rstPieceElec.Open("SELECT * FROM GrbCatalogueElec WHERE PIECE = '" & rstProjSoum.Fields("NumItem") & "'", g_connData, adOpenDynamic, adLockOptimistic)

  rstProjSoum.Fields("Desc_Fr") = rstPieceElec.Fields("DESC_FR")
  rstProjSoum.Fields("Desc_En") = rstPieceElec.Fields("DESC_EN")

  Call rstProjSoum.Update

  Call rstPieceElec.Close

  Call rstProjSoum.MoveNext
  Loop

  Set rstPieceElec = Nothing

10 Call rstProjSoum.Close
Set rstProjSoum = Nothing

Call RemplirListViewProjSoum(txtNoProjSoum.Text)

Exit Sub

Oups:

wOups "frmProjSoumElec", "UpdateDescription", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnuler_Click()

 On Error GoTo Oups
 
 fraPieceTrouve.Visible = False
 frafournisseur.Visible = False
 fraCommentaire.Visible = False
 fraDateRequise.Visible = False
 
 Screen.MousePointer = vbHourglass

 Call OuvrirProjSoum(False)
 
 'Remet en mode inactif
 Call AfficherControles(MODE_INACTIF)
 
 m_bEnregistrement = True
 
 'Affiche l'enregistrement qui était actif avant
 Call AfficherProjSoum(m_sAncienProjSoum)

 m_bEnregistrement = False
 
  m_bModeAjout = False
 
  Screen.MousePointer = vbDefault

  Exit Sub

Oups:

  wOups "frmProjSoumElec", "cmdAnnuler_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnulerCommentaire_Click()

 On Error GoTo Oups

 fraCommentaire.Visible = False

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "cmdAnnulerCommentaire_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnulerDateRequise_Click()

 On Error GoTo Oups

 fraDateRequise.Visible = False

 m_bMonthViewHasFocus = False

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "cmdAnnulerDateRequise_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnulerDateRequise_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
 On Error GoTo Oups

 If m_bMonthViewHasFocus = True Then
 Call cmdAnnulerDateRequise_Click
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "cmdAnnulerDateRequise_MouseUp", Err, Err.number, Err.Description
End Sub

Private Sub cmdExport_Click()

 On Error GoTo Oups

 Dim rstProjSoum As ADODB.Recordset

 Set rstProjSoum = New ADODB.Recordset

 Call g_connData.Execute("DELETE * FROM Grbimpression_listepiece")

 If m_eType = TYPE_PROJET Then
 Call rstProjSoum.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstProjSoum.Open("SELECT * FROM GrbSoumissionElec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
 End If

 Call ExporterListePieces(rstProjSoum)

 Call rstProjSoum.Close
  Set rstProjSoum = Nothing

  Exit Sub

Oups:

  wOups "frmProjSoumElec", "cmdExport_Click", Err, Err.number, Err.Description
End Sub

Private Sub ExporterListePieces(ByVal rstProjSoum As ADODB.Recordset)

 On Error GoTo Oups

 'Impression de la feuille de la liste des pièces
 Dim rstPiece As ADODB.Recordset
 Dim rstTemp As ADODB.Recordset
 Dim rstImpListePiece As ADODB.Recordset
 Dim iCompteurPiece As Integer
 Dim sSousSection As String
 Dim sSection As String
 Dim sNoProjet As String
 Dim sNoSoumission As String
 Dim bAjouterSection As Boolean
 Dim bAjouterSousSection As Boolean
  Dim bAjouterPiece As Boolean
  Dim xlsApp As Excel.Application
  Dim xlsWorkBook As Excel.Workbook
  Dim iCompteur As Integer
  Dim sSaveAsFileName As String

  Set rstPiece = New ADODB.Recordset
  Set rstTemp = New ADODB.Recordset
  Set rstImpListePiece = New ADODB.Recordset

10 iCompteurPiece = 1

Screen.MousePointer = vbHourglass

 'Ouverture du recordset
If m_eType = TYPE_PROJET Then
 sNoProjet = rstProjSoum.Fields("IDProjet")
 sNoSoumission = rstProjSoum.Fields("IDSoumission")

 Call rstPiece.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & sNoProjet & "' AND Type = 'E' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
Else
 sNoProjet = vbNullString
 sNoSoumission = rstProjSoum.Fields("IDSoumission")

 Call rstPiece.Open("SELECT * FROM GrbSoumission_Pieces WHERE IDSoumission = '" & sNoSoumission & "' AND Type = 'E' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
End If

Do While Not rstPiece.EOF
If rstPiece.Fields("Visible") = True Then
 bAjouterSection = True
 bAjouterSousSection = True
 bAjouterPiece = True

 rstImpListePiece.CursorLocation = adUseClient

 Call rstImpListePiece.Open("SELECT * FROM GrbImpression_ListePiece WHERE IDSection = '" & rstPiece.Fields("IDSection") & "'", g_connData, adOpenDynamic, adLockOptimistic)

 If Not rstImpListePiece.EOF Then
1  bAjouterSection = False

 Do While Not rstImpListePiece.EOF
 If rstImpListePiece.Fields("NomSousSection") = rstPiece.Fields("SousSection") Then
 bAjouterSousSection = False

 If rstPiece.Fields("NumItem") <> "Texte" And rstPiece.Fields("NumItem") <> "Text" Then
 If rstImpListePiece.Fields("NumItem") = rstPiece.Fields("NumItem") Then
 bAjouterPiece = False

 rstImpListePiece.Fields("Qté") = Replace(CDbl(rstImpListePiece.Fields("Qté")) + CDbl(rstPiece.Fields("Qté")), ".", ",")

 If Not IsNull(rstImpListePiece.Fields("ID")) Then
 If rstImpListePiece.Fields("ID") <> "" Then
 rstImpListePiece.Fields("ID") = rstImpListePiece.Fields("ID") & ", " & rstPiece.Fields("ID")
 Else
 rstImpListePiece.Fields("ID") = rstPiece.Fields("ID")
 End If
 Else
 rstImpListePiece.Fields("ID") = rstPiece.Fields("ID")
 End If

 Call rstImpListePiece.Update

 If rstImpListePiece.Fields("Qté") = 0 Then
 Call rstImpListePiece.Delete

 rstImpListePiece.Filter = "NomSousSection = '" & Replace(rstPiece.Fields("SousSection"), "'", "''") & "'"

 If rstImpListePiece.RecordCount = 1 Then
 Call rstImpListePiece.Delete

 rstImpListePiece.Filter = "IDSection = '" & rstPiece.Fields("IDSection") & "'"

 If rstImpListePiece.RecordCount = 1 Then
 Call rstImpListePiece.Delete
 End If
 End If

 rstImpListePiece.Filter = ""
 End If

 Exit Do
 End If
 Else
 Exit Do
 End If
 End If

 Call rstImpListePiece.MoveNext
 Loop
 End If

 If bAjouterSection = True Then
 If m_eLangage = ANGLAIS Then
 sSection = "NomSectionEN"
4 Else
4 sSection = "NomSectionFR"
4 End If

4 Call rstTemp.Open("SELECT " & sSection & " FROM GrbSoumProjSectionElec WHERE IDSection = " & rstPiece.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)

 'Ajoute la section dans la liste de pièces
4 Call rstImpListePiece.AddNew

4 rstImpListePiece.Fields("NoLigne") = iCompteurPiece
4 rstImpListePiece.Fields("IDSoumission") = sNoSoumission

4 If Not IsNull(rstTemp.Fields(sSection)) Then
4 rstImpListePiece.Fields("Section") = rstTemp.Fields(sSection)
4 Else
4 rstImpListePiece.Fields("Section") = " "
4  End If

4  rstImpListePiece.Fields("IDSection") = rstPiece.Fields("IDSection")

4  Call rstImpListePiece.Update

4  iCompteurPiece = iCompteurPiece + 1

4  Call rstTemp.Close
4  End If

4  If bAjouterSousSection = True Then
4  sSousSection = rstPiece.Fields("SousSection")

50 If sSousSection = S_PAS_SOUS_SECTION Then
 sSousSection = " "
 End If

 Call rstImpListePiece.AddNew

 rstImpListePiece.Fields("NoLigne") = iCompteurPiece
 rstImpListePiece.Fields("IDSoumission") = sNoSoumission
 rstImpListePiece.Fields("SousSection") = sSousSection
 rstImpListePiece.Fields("NomSousSection") = rstPiece.Fields("SousSection")
 rstImpListePiece.Fields("IDSection") = rstPiece.Fields("IDSection")

 Call rstImpListePiece.Update

 iCompteurPiece = iCompteurPiece + 1
 End If

5  If bAjouterPiece = True Then
 'Ajoute la pièce à la liste de pièces
5  Call rstImpListePiece.AddNew

5  rstImpListePiece.Fields("NoLigne") = iCompteurPiece
5  rstImpListePiece.Fields("IDSoumission") = sNoSoumission
5  rstImpListePiece.Fields("NumItem") = rstPiece.Fields("NumItem")
5  rstImpListePiece.Fields("Qté") = rstPiece.Fields("Qté")

5  If m_eLangage = ANGLAIS Then
5  rstImpListePiece.Fields("Description") = rstPiece.Fields("Desc_EN")
60 Else
  rstImpListePiece.Fields("Description") = rstPiece.Fields("Desc_FR")
  End If

  rstImpListePiece.Fields("Manufact") = rstPiece.Fields("Manufact")

  If m_eType = TYPE_PROJET Then
  rstImpListePiece.Fields("ID") = rstPiece.Fields("ID")
  End If

  rstImpListePiece.Fields("IDSection") = rstPiece.Fields("IDSection")
  rstImpListePiece.Fields("NomSousSection") = rstPiece.Fields("SousSection")

  Call rstImpListePiece.Update

  iCompteurPiece = iCompteurPiece + 1
  End If

6  Call rstImpListePiece.Close
6  End If

 'Prochaine enregistrement
6  Call rstPiece.MoveNext
6  Loop

 ''''''''''''''''''''''''''''''''''''''''''''''''''
 ' Rapport liste pièce, met dans l'ordre de ligne '
 ''''''''''''''''''''''''''''''''''''''''''''''''''
6  rstImpListePiece.CursorLocation = adUseClient

6  Call rstImpListePiece.Open("SELECT * FROM Grbimpression_Listepiece WHERE TRIM(IDSoumission) = '" & Trim$(sNoSoumission) & "' ORDER BY noligne", g_connData, adOpenDynamic, adLockOptimistic)

6  Set xlsApp = New Excel.Application

6  Set xlsWorkBook = xlsApp.Workbooks.Add

70 xlsApp.range("A1") = "Liste de matériel ( " & txtNoProjSoum.Text & " )"
70 xlsApp.range("A1").Font.Bold = True
  xlsApp.range("A1").Font.Underline = xlUnderlineStyleSingle
  xlsApp.range("A1").HorizontalAlignment = xlCenter
  xlsApp.range("A1").Font.SIZE = 14

  Call xlsApp.range("A1:E1").Merge

  xlsApp.range("A4") = "Qté"
  xlsApp.range("A4").Font.Bold = True
  xlsApp.range("A4").HorizontalAlignment = xlCenter

  xlsApp.range("B4") = "No. Item"
  xlsApp.range("B4").Font.Bold = True
  xlsApp.range("B4").HorizontalAlignment = xlCenter

   xlsApp.range("C4") = "Description"
   xlsApp.range("C4").Font.Bold = True
7  xlsApp.range("C4").HorizontalAlignment = xlCenter

7  xlsApp.range("D4") = "Manufacturier"
7  xlsApp.range("D4").Font.Bold = True
7  xlsApp.range("D4").HorizontalAlignment = xlCenter

7  xlsApp.range("E4") = "#ID"
7  xlsApp.range("E4").Font.Bold = True
80 xlsApp.range("E4").HorizontalAlignment = xlCenter

80 xlsApp.range("A4:E4").Borders(xlEdgeBottom).LineStyle = xlContinuous
  xlsApp.range("A4:E4").Borders(xlEdgeBottom).Weight = xlMedium
  xlsApp.range("A4:E4").Borders(xlEdgeBottom).ColorIndex = xlAutomatic

  xlsApp.range("A4:E4").Borders(xlInsideVertical).LineStyle = xlContinuous
  xlsApp.range("A4:E4").Borders(xlInsideVertical).Weight = xlMedium
  xlsApp.range("A4:E4").Borders(xlInsideVertical).ColorIndex = xlAutomatic

  iCompteur = 5

  Do While Not rstImpListePiece.EOF
  xlsApp.range("A" & iCompteur) = rstImpListePiece.Fields("Qté")

  If IsNull(rstImpListePiece.Fields("Section")) Then
  xlsApp.range("B" & iCompteur) = rstImpListePiece.Fields("NumItem")
   Else
   xlsApp.range("B" & iCompteur) = rstImpListePiece.Fields("Section")
   xlsApp.range("B" & iCompteur).Font.Bold = True
   End If

8  If IsNull(rstImpListePiece.Fields("SousSection")) Then
8  xlsApp.range("C" & iCompteur) = rstImpListePiece.Fields("Description")
8  Else
8  xlsApp.range("C" & iCompteur) = rstImpListePiece.Fields("SousSection")
90 xlsApp.range("C" & iCompteur).Font.Bold = True
  End If

  xlsApp.range("D" & iCompteur) = rstImpListePiece.Fields("Manufact")
  xlsApp.range("E" & iCompteur) = rstImpListePiece.Fields("ID")

  xlsApp.range("A" & iCompteur & ":E" & iCompteur).Borders(xlEdgeBottom).LineStyle = xlContinuous
  xlsApp.range("A" & iCompteur & ":E" & iCompteur).Borders(xlEdgeBottom).Weight = xlThin
  xlsApp.range("A" & iCompteur & ":E" & iCompteur).Borders(xlEdgeBottom).ColorIndex = xlAutomatic

  xlsApp.range("A" & iCompteur & ":E" & iCompteur).Borders(xlInsideVertical).LineStyle = xlContinuous
  xlsApp.range("A" & iCompteur & ":E" & iCompteur).Borders(xlInsideVertical).Weight = xlThin
  xlsApp.range("A" & iCompteur & ":E" & iCompteur).Borders(xlInsideVertical).ColorIndex = xlAutomatic

  Call rstImpListePiece.MoveNext

  iCompteur = iCompteur + 1
 Loop

   iCompteur = iCompteur - 1

 xlsApp.range("A4:E" & iCompteur).Borders(xlEdgeBottom).LineStyle = xlContinuous
   xlsApp.range("A4:E" & iCompteur).Borders(xlEdgeBottom).Weight = xlMedium
 xlsApp.range("A4:E" & iCompteur).Borders(xlEdgeBottom).ColorIndex = xlAutomatic

   xlsApp.range("A4:E" & iCompteur).Borders(xlEdgeTop).LineStyle = xlContinuous
 xlsApp.range("A4:E" & iCompteur).Borders(xlEdgeTop).Weight = xlMedium
9  xlsApp.range("A4:E" & iCompteur).Borders(xlEdgeTop).ColorIndex = xlAutomatic

 xlsApp.range("A4:E" & iCompteur).Borders(xlEdgeLeft).LineStyle = xlContinuous
100 xlsApp.range("A4:E" & iCompteur).Borders(xlEdgeLeft).Weight = xlMedium
10 xlsApp.range("A4:E" & iCompteur).Borders(xlEdgeLeft).ColorIndex = xlAutomatic

10 xlsApp.range("A4:E" & iCompteur).Borders(xlEdgeRight).LineStyle = xlContinuous
xlsApp.range("A4:E" & iCompteur).Borders(xlEdgeRight).Weight = xlMedium
10 xlsApp.range("A4:E" & iCompteur).Borders(xlEdgeRight).ColorIndex = xlAutomatic

Call xlsApp.Columns("A:A").EntireColumn.AutoFit
10 Call xlsApp.Columns("B:B").EntireColumn.AutoFit
Call xlsApp.Columns("C:C").EntireColumn.AutoFit
10 Call xlsApp.Columns("D:D").EntireColumn.AutoFit
Call xlsApp.Columns("E:E").EntireColumn.AutoFit

10 Call rstImpListePiece.Close
10  Set rstImpListePiece = Nothing

10  Screen.MousePointer = vbDefault

10  sSaveAsFileName = xlsApp.GetSaveAsFilename(txtNoProjSoum.Text & ".xls", "Fichiers Excel (*.xlx), *.xls")

10  If sSaveAsFileName <> "Faux" Then
10  Call xlsWorkBook.SaveAs(sSaveAsFileName)
10  End If

10  xlsWorkBook.Saved = True

10  Call xlsWorkBook.Close

110 Set xlsWorkBook = Nothing

110 Call xlsApp.Quit

11 Set xlsApp = Nothing

11 Set rstTemp = Nothing

11 Exit Sub

Oups:

11 wOups "frmProjSoumElec", "ExporterListePieces", Err, Err.number, Err.Description
End Sub

Private Sub cmdOKCommentaire_Click()

 On Error GoTo Oups

 lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_COMMENTAIRE) = txtcommentaire.Text

 fraCommentaire.Visible = False

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "cmdOKCommentaire_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdOKDateRequise_Click()

 On Error GoTo Oups

 Dim datDate As Date

 datDate = DateSerial(mvwDateRequise.Year, mvwDateRequise.Month, mvwDateRequise.Day)

 lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_DATE_REQUISE) = ConvertDate(datDate)

 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_DATE_REQUISE).ForeColor = COLOR_ORANGE

 fraDateRequise.Visible = False

 m_bMonthViewHasFocus = False

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "cmdOKDateRequise_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnulerFRS_Click()

 On Error GoTo Oups

 frafournisseur.Visible = False

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "cmdAnnulerFRS_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnulerPieceTrouve_Click()

 On Error GoTo Oups

 fraPieceTrouve.Visible = False

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "cmdAnnulerPieceTrouve", Err, Err.number, Err.Description
End Sub

Private Sub cmdBonCommande_Click()

 On Error GoTo Oups

 Dim rstProjSoum As ADODB.Recordset
 Dim sUser As String

 If Right$(txtNoProjSoum.Text, 2) = "99" Then
 Call MsgBox("Vous ne pouvez pas commander de pièce à partir de ce projet!", vbOKOnly, "Erreur")

 Exit Sub
 End If

 Set rstProjSoum = New ADODB.Recordset

 Call rstProjSoum.Open("SELECT * FROM GrbProjSoum WHERE IDProjSoum = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

 If rstProjSoum.Fields("Ouvert") = True And rstProjSoum.Fields("Verrouillé") = False Then
 If VerifierSiOuvert(sUser) = False Then
  If lvwSoumission.ListItems.count > 0 Then
  Call frmChoixBonCommande.Afficher(txtNoProjSoum.Text, Me, m_eLangage)
  Else
  Call MsgBox("Il n'y a pas de pièces à commander pour ce projet!", vbOKOnly, "Erreur")
  End If
  Else
  Call MsgBox("Ce projet est en modification par " & sUser & "!", vbOKOnly, "Erreur")
  End If
10 Else
1 If rstProjSoum.Fields("Ouvert") = False Then
 Call MsgBox("Ce projet est fermé!", vbOKOnly, "Erreur")
 Else
 Call MsgBox("Ce projet est verrouillé!", vbOKOnly, "Erreur")
 End If
End If

Call rstProjSoum.Close
Set rstProjSoum = Nothing

If m_bModifFournisseurBC = True Then
 Call cmbProjSoum_Click
End If

1  Exit Sub

Oups:

wOups "frmProjSoumElec", "cmdBonCommande_Click", Err, Err.number, Err.Description
End Sub

Public Sub Commande()

 On Error GoTo Oups
 
 'Change la valeur du champs "Commandé" pour la pièce à true
 Dim rstProjet As ADODB.Recordset
 Dim rstPiece As ADODB.Recordset
 Dim rstBCPiece As ADODB.Recordset
 Dim rstBC As ADODB.Recordset
 Dim rstFRS As ADODB.Recordset
 Dim iIDFRS As Integer
 Dim sFRS As String
 Dim sNoBC As String
 Dim sWherePiece As String
 Dim sWhereNoLigne As String
  Dim sWhere As String
  Dim sDateRequise As String
  Dim sNoLigne As String
  Dim bPremier As Boolean
  Dim bPremierNoLigne As Boolean

  Set rstProjet = New ADODB.Recordset

  Call rstProjet.Open("SELECT ProchaineCommande FROM GrbProjetElec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

  If Not IsNull(rstProjet.Fields("ProchaineCommande")) Then
rstProjet.Fields("ProchaineCommande") = rstProjet.Fields("ProchaineCommande") + 1

1 Call rstProjet.Update
End If

Call rstProjet.Close
Set rstProjet = Nothing

sFRS = DR_Commande.Sections("Section2").Controls("lblFournisseur").Caption
sNoBC = DR_Commande.Sections("Section2").Controls("lblNoBC").Caption

Set rstBC = New ADODB.Recordset
Set rstFRS = New ADODB.Recordset
Set rstPiece = New ADODB.Recordset
Set rstBCPiece = New ADODB.Recordset

Call rstBC.Open("SELECT * FROM GrbBonsCommandes WHERE NoBonCommande = '" & sNoBC & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
1  Do While Not rstBC.EOF
 Call rstFRS.Open("SELECT IDFRS, NomFournisseur FROM GrbFournisseur WHERE IDFRS = " & rstBC.Fields("NoFournisseur"), g_connData, adOpenDynamic, adLockOptimistic)

 If rstFRS.Fields("NomFournisseur") = sFRS Then
 iIDFRS = rstFRS.Fields("IDFRS")

 sDateRequise = rstBC.Fields("DateRequise")

 Call rstFRS.Close

 Exit Do
1  End If

 Call rstFRS.Close

 Call rstBC.MoveNext
Loop

Call rstBC.Close
Set rstBC = Nothing

Set rstFRS = Nothing
 
 'Ouverture du recordset du Bon de commande pour savoir quelles pièces
 'ont été commandées
Call rstBCPiece.Open("SELECT NoItem, NuméroLigne FROM GrbBonsCommandes_Pieces WHERE NoFournisseur = " & iIDFRS & " AND NoBonCommande = '" & sNoBC & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Tant que ce n'est pas la fin des enregistrements
sWhere = "(IDProjet = '" & txtNoProjSoum.Text & "')"
 
sWherePiece = "NumItem In ("
sWhereNoLigne = "NuméroLigne In ("
 
bPremier = True
 
Do While Not rstBCPiece.EOF
If Not IsNull(rstBCPiece.Fields("NoItem")) Then
 sNoLigne = rstBCPiece.Fields("NuméroLigne")

 If bPremier = True Then
 If InStr(1, sNoLigne, ",") = 0 Then
 sWherePiece = sWherePiece & "'" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
 sWhereNoLigne = sWhereNoLigne & rstBCPiece.Fields("NuméroLigne")
 Else
 bPremierNoLigne = True

 Do While InStr(1, sNoLigne, ",") > 0
 If bPremierNoLigne = True Then
 sWherePiece = sWherePiece & "'" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
 sWhereNoLigne = sWhereNoLigne & Left$(sNoLigne, InStr(1, sNoLigne, ",") - 1)

 bPremierNoLigne = False
 Else
 sWherePiece = sWherePiece & ", '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
 sWhereNoLigne = sWhereNoLigne & ", " & Left$(sNoLigne, InStr(1, sNoLigne, ",") - 1)
 End If

 sNoLigne = Right$(sNoLigne, Len(sNoLigne) - (InStr(1, sNoLigne, ",") + 1))
 Loop

 If Trim$(sNoLigne) <> "" Then
 sWherePiece = sWherePiece & ", '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
 sWhereNoLigne = sWhereNoLigne & ", " & sNoLigne
 End If
 End If

 bPremier = False
 Else
 If InStr(1, sNoLigne, ",") = 0 Then
 sWherePiece = sWherePiece & ", '" & rstBCPiece.Fields("NoItem") & "'"
 sWhereNoLigne = sWhereNoLigne & ", " & rstBCPiece.Fields("NuméroLigne")
4 Else
4 Do While InStr(1, sNoLigne, ",") > 0
4 sWherePiece = sWherePiece & ", '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
4 sWhereNoLigne = sWhereNoLigne & ", " & Left$(sNoLigne, InStr(1, sNoLigne, ",") - 1)

4 sNoLigne = Right$(sNoLigne, Len(sNoLigne) - (InStr(1, sNoLigne, ",") + 1))
4 Loop

4 If Trim$(sNoLigne) <> "" Then
4 sWherePiece = sWherePiece & ", '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
4 sWhereNoLigne = sWhereNoLigne & ", " & sNoLigne
4 End If
4 End If
4  End If
4  End If
 
4  Call rstBCPiece.MoveNext
4  Loop

4  sWherePiece = sWherePiece & ")"
4  sWhereNoLigne = sWhereNoLigne & ")"

4  sWhere = sWhere & " AND " & sWherePiece & " AND " & sWhereNoLigne
 
4  Call rstBCPiece.Close
50 Set rstBCPiece = Nothing
 
50 Call rstPiece.Open("SELECT * FROM GrbProjet_Pieces WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
 
 Do While Not rstPiece.EOF
 rstPiece.Fields("Commandé") = True

 rstPiece.Fields("DateCommande") = ConvertDate(Date)

 rstPiece.Fields("DateRequise") = sDateRequise

 rstPiece.Fields("NomCommande") = g_sEmploye

 rstPiece.Fields("NoSéquentiel") = Right$(sNoBC, 3)
 
 Call rstPiece.Update
 
 Call rstPiece.MoveNext
 Loop
 
 Call rstPiece.Close
5  Set rstPiece = Nothing
 
5  Call RemplirListViewProjSoum(txtNoProjSoum.Text)

5  Exit Sub

Oups:

5  wOups "frmProjSoumElec", "Commande", Err, Err.number, Err.Description
End Sub

Private Sub cmdCatalogue_Click()

 On Error GoTo Oups

 'Pour ouvrir le catalogue électrique
 Screen.MousePointer = vbHourglass

 Call FrmCatalogueElec.AfficherForm(cmbPieces.Text, "", "")

 Screen.MousePointer = vbDefault

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "cmdCatalogue_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdForfait_Click()

 On Error GoTo Oups

 Dim sMontant As String

 sMontant = InputBox("Quel est le montant du forfait?")

 If Trim$(sMontant) <> "" Then
 sMontant = Replace(sMontant, ".", ",")

 If IsNumeric(sMontant) Then
 txtForfait.Text = Conversion(sMontant, MODE_ARGENT)

 lblForfaitInitiale.Caption = g_sInitiale
 Else
 Call MsgBox("Montant non numérique!", vbOKOnly, "Erreur")
 End If
  End If

  Exit Sub

Oups:

  wOups "frmProjSoumElec", "cmdForfait_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdMauvaisPrix_Click()

 On Error GoTo Oups

 Call MauvaisPrix

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "cmdMauvaisPrix_Click", Err, Err.number, Err.Description
End Sub

Private Sub MauvaisPrix()

 On Error GoTo Oups

 Dim iCompteur As Integer

 If lvwSoumission.ListItems.count > 0 Then
 If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PROFIT).Tag <> "EXTRA" Then
 'Si ce n'est pas une section
 If lvwSoumission.SelectedItem.Tag <> vbNullString Then
 'Si ce n'est pas une sous-section
 If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) <> vbNullString Then
 'Si ce n'est pas du texte
 If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) <> "Texte" And lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) <> "Text" Then
 'Si la quantité est plus grande que 0
 If CDbl(Replace(lvwSoumission.SelectedItem.Text, "*", vbNullString)) > 0 Then
 Call ViderChamps_frs

 Call RemplirComboFournisseur

 For iCompteur = 0 To cmbfrs.ListCount - 1
  If cmbfrs.ItemData(iCompteur) = lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_DISTRIB).Tag Then
  cmbfrs.ListIndex = iCompteur

  Exit For
  End If
  Next

  cmbfrs.Locked = True

  fraPrixPiece.Tag = lvwSoumission.SelectedItem.Index

  m_bMauvaisPrix = True

 fraPrixPiece.Visible = True

 Call txtPrixList.SetFocus
 Else
 Call MsgBox("La quantité est déjà dans le négatif!", vbOKOnly, "Erreur")
 End If
 End If
 End If
 End If
 Else
 Call MsgBox("Cette commande doit être faite dans le projet " & lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PROVENANCE), vbOKOnly, "Erreur")
 End If
End If

1  Exit Sub

Oups:

wOups "frmProjSoumElec", "MauvaisPrix", Err, Err.number, Err.Description
End Sub

Private Sub cmdMaterielInutile_Click()

 On Error GoTo Oups

 Call MaterielInutile

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "cmdMaterielInutile_Click", Err, Err.number, Err.Description
End Sub

Private Sub MaterielInutile()

 On Error GoTo Oups

 Dim itmProjet As ListItem

 If lvwSoumission.ListItems.count > 0 Then
 Set itmProjet = lvwSoumission.SelectedItem

 If itmProjet.ListSubItems(I_COL_SOUM_PIECE).ForeColor <> COLOR_ROSE And itmProjet.ListSubItems(I_COL_SOUM_PIECE).ForeColor <> COLOR_BLEU Then
 'Si ce n'est pas une section
 If itmProjet.Tag <> vbNullString Then
 'Si ce n'est pas une sous-section
 If itmProjet.SubItems(I_COL_SOUM_PIECE) <> vbNullString Then
 'Si ce n'est pas du texte
 If itmProjet.SubItems(I_COL_SOUM_PIECE) <> "Texte" And itmProjet.SubItems(I_COL_SOUM_PIECE) <> "Text" Then
 'Si la quantité est plus grande que 0
 If CDbl(Replace(itmProjet.Text, "*", vbNullString)) > 0 Then
 m_bPieceInutile = True
 m_bRecherchePiece = False
  m_bChangementFRS = False

  Call AfficherListeFournisseurs

  If lvwfournisseur.ListItems.count = 0 Then
  Call MsgBox("Il n'y a aucun fournisseur pour cette pièce!", vbOKOnly, "Erreur")
 
  Exit Sub
  Else
  frafournisseur.Visible = True
  End If
 Else
 Call MsgBox("La quantité est déjà dans le négatif!", vbOKOnly, "Erreur")
 End If
 End If
 End If
 End If
 Else
 Call MsgBox("Cette commande doit être faite dans le projet " & itmProjet.SubItems(I_COL_SOUM_PROVENANCE), vbOKOnly, "Erreur")
 End If
End If

Exit Sub

Oups:

wOups "frmProjSoumElec", "MaterielInutile", Err, Err.number, Err.Description
End Sub

Private Sub cmdCopier_Click()

 On Error GoTo Oups

 Dim rstProjSoum As ADODB.Recordset
 Dim sNoProjSoum As String
 Dim sUser As String
 Dim bExiste As Boolean
 Dim bVariables As Boolean
 Dim bTauxHoraire As Boolean
 Dim bPrixPieces As Boolean
 Dim bNoValide As Boolean
 
 'Si le combo n'est pas vide
 If cmbProjSoum.ListCount > 0 Then
 If VerifierSiOuvert(sUser) = False Then
 'Demande du numéro
  sNoProjSoum = InputBox("Quel est le numéro de la soumission?")
 
  If Trim$(sNoProjSoum) <> vbNullString Then
  Screen.MousePointer = vbHourglass

  bNoValide = True

  If ValiderFormatNumeroProjSoum(sNoProjSoum) = False Then
  bNoValide = False
  End If

  If bNoValide = True Then
 If ValiderFormatElectrique(sNoProjSoum) = False Then
 bNoValide = False
 End If
 End If

 If bNoValide = True Then
 If ValiderFormatSoumission(sNoProjSoum) = False Then
 bNoValide = False
 End If
 End If

 If bNoValide = False Then
 Screen.MousePointer = vbDefault

 Exit Sub
 End If

 sNoProjSoum = UCase(sNoProjSoum)
 
 Set rstProjSoum = New ADODB.Recordset
 
 Call rstProjSoum.Open("SELECT * FROM GrbSoumissionElec WHERE IDSoumission = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 If rstProjSoum.EOF Then
 bExiste = False
 Else
1  bExiste = True

 Call MsgBox("Ce numéro existe dans les soumissions électriques!", vbOKOnly, "Erreur")
 End If

 Call rstProjSoum.Close

 If bExiste = False Then
 Call rstProjSoum.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 If rstProjSoum.EOF Then
 bExiste = False
 Else
 bExiste = True

 Call MsgBox("Ce numéro existe dans les projets électriques!", vbOKOnly, "Erreur")
 End If

 Call rstProjSoum.Close
 End If

 If bExiste = False Then
 Call rstProjSoum.Open("SELECT * FROM GrbSoumissionMec WHERE IDSoumission = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)

 If rstProjSoum.EOF Then
 bExiste = False
 Else
 bExiste = True
 
 Call MsgBox("Ce numéro existe dans les soumissions mécaniques!", vbOKOnly, "Erreur")
 End If

 Call rstProjSoum.Close
 End If
 
 If bExiste = False Then
 Call rstProjSoum.Open("SELECT * FROM GrbProjetMec WHERE IDProjet = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 If rstProjSoum.EOF Then
 bExiste = False
 Else
 bExiste = True
 
 Call MsgBox("Ce numéro existe dans les projets mécaniques!", vbOKOnly, "Erreur")
 End If

 Call rstProjSoum.Close
 End If
 
 'S'il n'existe pas, on l'ajoute
 If bExiste = False Then
 Call rstProjSoum.Open("SELECT * FROM GrbProjSoum WHERE IDProjSoum = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)

 If Not rstProjSoum.EOF Then
 If rstProjSoum.Fields("Ouvert") = False Then
 Call MsgBox("Ce numéro est fermé!", vbOKOnly, "Erreur")
 
 Call rstProjSoum.Close
 Set rstProjSoum = Nothing
 
 Screen.MousePointer = vbDefault

4 Exit Sub
4 End If
4 End If

4 Call rstProjSoum.Close
4 Set rstProjSoum = Nothing

4 If MsgBox("Voulez-vous mettre à jour les variables systèmes?" & vbNewLine & _
 "- % Profit" & vbNewLine & _
 "- % Commission" & vbNewLine & _
 "- % Imprévu" & vbNewLine & _
 "- $ Pages manuel", vbYesNo) = vbYes Then
4 bVariables = True
4 Else
4 bVariables = False
4 End If

4 If MsgBox("Voulez-vous mettre à jour les taux horaires?", vbYesNo) = vbYes Then
4  bTauxHoraire = True
4  Else
4  bTauxHoraire = False
4  End If

4  If MsgBox("Voulez-vous mettre à jour le prix des pièces?", vbYesNo) = vbYes Then
4  bPrixPieces = True
4  Else
4  bPrixPieces = False
50 End If

 m_bModeAjout = True
 m_bModeAffichage = False
 
 m_bTempsDejaOuvert = True
 
 If bVariables = True Then
 'On ré-initialise les variables
 Call InitialiserVariables(sNoProjSoum)
 End If

 If bTauxHoraire = True Then
 Call InitialiserNouveauxTaux
 End If
 
 'Rapetisse le listview de la soumission pour afficher le lvwPiece
 lvwSoumission.Height = lvwSoumission.Height * 0.49
 lvwSoumission.Top = lvwPieces.Top + lvwPieces.Height + (lvwSoumission.Height * 0.02)
 
 'On met en mode modif
5  Call AfficherControles(MODE_AJOUT_MODIF)
 
5  If bPrixPieces = True Then
 'On recalcul le prix des pièces
5  Call UpdatePieces
5  End If
 
5  Call UpdateOrdre
 
5  If bVariables = True Or bTauxHoraire = True Or bPrixPieces = True Then
 'On recalcul le prix total
5  Call CalculerPrix
5  End If

60 Call BarrerChamps(False)
 
  txtNoProjSoum.Text = sNoProjSoum
  txtNoSoumission.Text = vbNullString
  End If
 
  Screen.MousePointer = vbDefault
  End If
  Else
  If m_eType = TYPE_PROJET Then
  Call MsgBox("Ce projet est en modification par " & sUser & "!", vbOKOnly, "Erreur")
  Else
  Call MsgBox("Cette soumission est en modification par " & sUser & "!", vbOKOnly, "Erreur")
  End If
6  End If
6  End If

6  Exit Sub

Oups:

6  wOups "frmProjSoumElec", "cmdCopier_Click", Err, Err.number, Err.Description
End Sub

Private Sub UpdateOrdre()

 On Error GoTo Oups

 'Cette procédure sert à changer l'ordre des sections dans la soumission
 Dim rstOrdre As ADODB.Recordset
 Dim rstCount As ADODB.Recordset
 Dim rstSection As ADODB.Recordset
 Dim iCompteur As Integer
 Dim iCompteurAs Integer
 Dim iIndexCopie As Integer
 Dim iSection As Integer
 Dim iIndex As Integer
 Dim iNbreSection As Integer
 Dim bPremier As Boolean
  Dim itmProjSoum As ListItem
  Dim sSection As String
 
  Set rstOrdre = New ADODB.Recordset
 
 'Boucle pour changer la valeur de l'ordre dans le ListItem
  For iCompteur = 1 To lvwSoumission.ListItems.count
 'Si ce n'est pas une section
  If lvwSoumission.ListItems(iCompteur).Tag <> vbNullString Then
  Call rstOrdre.Open("SELECT Ordre FROM GrbSoumProjSectionElec WHERE IDSection = " & lvwSoumission.ListItems(iCompteur).Tag, g_connData, adOpenDynamic, adLockOptimistic)
 
  lvwSoumission.ListItems(iCompteur).ListSubItems(I_COL_SOUM_MANUFACT).Tag = rstOrdre.Fields("Ordre")
 
  Call rstOrdre.Close
End If
Next

Set rstOrdre = Nothing
 
Set rstCount = New ADODB.Recordset
 
Call rstCount.Open("SELECT COUNT(IDSection) as NbreSection FROM GrbSoumProjSectionElec", g_connData, adOpenDynamic, adLockOptimistic)

iNbreSection = rstCount.Fields("NbreSection")

Call rstCount.Close
Set rstCount = Nothing

 'Il faut enlever les sections car ils n'ont pas d'ordre et il ne font
 'que nuire
For iCompteur = 1 To lvwSoumission.ListItems.count
 'Si c'est une section
 If lvwSoumission.ListItems(iCompteur - iSection).Tag = vbNullString Then
 'On l'enlève
 Call lvwSoumission.ListItems.Remove(iCompteur - iSection)

 iSection = iSection + 1
End If
Next

 iIndex = 1

Set rstSection = New ADODB.Recordset

 'Boucle pour replacer le ListItem à la bonne place
 For iCompteur = 1 To iNbreSection
 bPremier = True

 iCompteur2 = iIndex

1  Do While iCompteur2 <= lvwSoumission.ListItems.count
 'Si le tag est la premiere ordre
 If lvwSoumission.ListItems(iCompteur2).ListSubItems(I_COL_SOUM_MANUFACT).Tag = iCompteur Then
 'Si la première fois qu'on trouve cette ordre
 If bPremier = True Then
 'on ajoute la section
 Set itmProjSoum = lvwSoumission.ListItems.Add(iIndex)

 Call ValeurParDefaut(itmProjSoum)
 
 If m_eLangage = ANGLAIS Then
 sSection = "NomSectionEN"
 Else
 sSection = "NomSectionFR"
 End If

 Call rstSection.Open("SELECT " & sSection & " FROM GrbSoumProjSectionElec WHERE IDSection = " & lvwSoumission.ListItems(iCompteur2 + 1).Tag, g_connData, adOpenDynamic, adLockOptimistic)

 If Not IsNull(rstSection.Fields(sSection)) Then
 itmProjSoum.SubItems(I_COL_SOUM_PIECE) = rstSection.Fields(sSection)
 Else
 itmProjSoum.SubItems(I_COL_SOUM_PIECE) = vbNullString
 End If
 
 itmProjSoum.ListSubItems(I_COL_SOUM_PIECE).Bold = True

 Call rstSection.Close

 iIndex = iIndex + 1
 iCompteur2 = iCompteur2 + 1
 
 bPremier = False
 End If

 'On ajoute la pièce
 Set itmProjSoum = lvwSoumission.ListItems.Add(iIndex)

 iIndexCopie = iCompteur2 + 1
 
 itmProjSoum.Checked = lvwSoumission.ListItems(iIndexCopie).Checked
 
 itmProjSoum.Text = lvwSoumission.ListItems(iIndexCopie).Text
 
 itmProjSoum.ForeColor = lvwSoumission.ListItems(iIndexCopie).ForeColor
 
 itmProjSoum.Tag = lvwSoumission.ListItems(iIndexCopie).Tag

 itmProjSoum.Bold = lvwSoumission.ListItems(iIndexCopie).Bold

 itmProjSoum.SubItems(I_COL_SOUM_PIECE) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_PIECE)
 itmProjSoum.ListSubItems(I_COL_SOUM_PIECE).Tag = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_PIECE).Tag
 
 itmProjSoum.ListSubItems(I_COL_SOUM_PIECE).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_PIECE).ForeColor

 itmProjSoum.ListSubItems(I_COL_SOUM_PIECE).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_PIECE).Bold

 itmProjSoum.SubItems(I_COL_SOUM_DESCR) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_DESCR)
 itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_DESCR).Tag
 
 itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_DESCR).ForeColor

 itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_DESCR).Bold = True

 itmProjSoum.SubItems(I_COL_SOUM_MANUFACT) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_MANUFACT)
 itmProjSoum.ListSubItems(I_COL_SOUM_MANUFACT).Tag = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_MANUFACT).Tag
 
 itmProjSoum.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_MANUFACT).ForeColor

 itmProjSoum.ListSubItems(I_COL_SOUM_MANUFACT).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_MANUFACT).Bold

 itmProjSoum.SubItems(I_COL_SOUM_PRIX_LIST) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_PRIX_LIST)
4 itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_PRIX_LIST).Tag

4 itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor

4 itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_PRIX_LIST).Bold

4 itmProjSoum.SubItems(I_COL_SOUM_ESCOMPTE) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_ESCOMPTE)

4 itmProjSoum.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor

4 itmProjSoum.ListSubItems(I_COL_SOUM_ESCOMPTE).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_ESCOMPTE).Bold

4 itmProjSoum.SubItems(I_COL_SOUM_PRIX_NET) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_PRIX_NET)

4 itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor

4 itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_NET).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_PRIX_NET).Bold

4 itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_NET).Tag = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_PRIX_NET).Tag

4 itmProjSoum.SubItems(I_COL_SOUM_TEMPS) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_TEMPS)
 
4  itmProjSoum.ListSubItems(I_COL_SOUM_TEMPS).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_TEMPS).ForeColor

4  itmProjSoum.ListSubItems(I_COL_SOUM_TEMPS).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_TEMPS).Bold

4  itmProjSoum.SubItems(I_COL_SOUM_MONTAGE) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_MONTAGE)
 
4  itmProjSoum.ListSubItems(I_COL_SOUM_MONTAGE).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_MONTAGE).ForeColor

4  itmProjSoum.ListSubItems(I_COL_SOUM_MONTAGE).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_MONTAGE).Bold

4  itmProjSoum.SubItems(I_COL_SOUM_TOTAL) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_TOTAL)
 
4  itmProjSoum.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_TOTAL).ForeColor

4  itmProjSoum.ListSubItems(I_COL_SOUM_TOTAL).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_TOTAL).Bold

50 itmProjSoum.ListSubItems(I_COL_SOUM_TOTAL).Tag = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_TOTAL).Tag

 itmProjSoum.SubItems(I_COL_SOUM_PROFIT) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_PROFIT)
 
 itmProjSoum.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_PROFIT).ForeColor

 itmProjSoum.ListSubItems(I_COL_SOUM_PROFIT).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_PROFIT).Bold

 itmProjSoum.ListSubItems(I_COL_SOUM_PROFIT).Tag = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_PROFIT).Tag

 itmProjSoum.SubItems(I_COL_SOUM_DISTRIB) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_DISTRIB)
 itmProjSoum.ListSubItems(I_COL_SOUM_DISTRIB).Tag = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_DISTRIB).Tag
 
 itmProjSoum.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_DISTRIB).ForeColor

 itmProjSoum.ListSubItems(I_COL_SOUM_DISTRIB).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_DISTRIB).Bold

 If lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_COMMENTAIRE) = "" Then
 itmProjSoum.SubItems(I_COL_SOUM_COMMENTAIRE) = " "
 Else
5  itmProjSoum.SubItems(I_COL_SOUM_COMMENTAIRE) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_COMMENTAIRE)
5  itmProjSoum.ListSubItems(I_COL_SOUM_COMMENTAIRE).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_COMMENTAIRE).ForeColor
5  itmProjSoum.ListSubItems(I_COL_SOUM_COMMENTAIRE).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_COMMENTAIRE).Bold
5  End If

5  If m_eType = TYPE_PROJET Then
5  If itmProjSoum.SubItems(I_COL_SOUM_PIECE) <> "" And itmProjSoum.SubItems(I_COL_SOUM_PIECE) <> "Texte" And itmProjSoum.SubItems(I_COL_SOUM_PIECE) <> "Text" Then
5  itmProjSoum.SubItems(I_COL_SOUM_ID) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_ID)

5  itmProjSoum.ListSubItems(I_COL_SOUM_ID).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_ID).ForeColor

60 itmProjSoum.ListSubItems(I_COL_SOUM_ID).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_ID).Bold

  itmProjSoum.SubItems(I_COL_SOUM_FACTURATION) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_FACTURATION)

  If lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_FACTURATION) = "" Then
  itmProjSoum.ListSubItems(I_COL_SOUM_FACTURATION).Tag = ""
  Else
  itmProjSoum.ListSubItems(I_COL_SOUM_FACTURATION).Tag = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_FACTURATION).Tag
  End If

  itmProjSoum.SubItems(I_COL_SOUM_DATE_COMMANDE) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_DATE_COMMANDE)

  If itmProjSoum.SubItems(I_COL_SOUM_DATE_COMMANDE) <> "" Then
  itmProjSoum.ListSubItems(I_COL_SOUM_DATE_COMMANDE).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_DATE_COMMANDE).ForeColor
  End If

  itmProjSoum.ListSubItems(I_COL_SOUM_DATE_COMMANDE).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_DATE_COMMANDE).Bold

6  itmProjSoum.ListSubItems(I_COL_SOUM_DATE_COMMANDE).Tag = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_DATE_COMMANDE).Tag

6  itmProjSoum.SubItems(I_COL_SOUM_DATE_REQUISE) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_DATE_REQUISE)

6  itmProjSoum.ListSubItems(I_COL_SOUM_DATE_REQUISE).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_DATE_REQUISE).ForeColor

6  itmProjSoum.ListSubItems(I_COL_SOUM_DATE_REQUISE).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_DATE_REQUISE).Bold

6  itmProjSoum.ListSubItems(I_COL_SOUM_DATE_REQUISE).Tag = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_DATE_REQUISE).Tag

6  itmProjSoum.SubItems(I_COL_SOUM_NOM_COMMANDE) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_NOM_COMMANDE)

6  itmProjSoum.ListSubItems(I_COL_SOUM_NOM_COMMANDE).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_NOM_COMMANDE).ForeColor

6  itmProjSoum.ListSubItems(I_COL_SOUM_NOM_COMMANDE).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_NOM_COMMANDE).Bold

70 itmProjSoum.SubItems(I_COL_SOUM_NO_SEQUENTIEL) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_NO_SEQUENTIEL)

  itmProjSoum.ListSubItems(I_COL_SOUM_NO_SEQUENTIEL).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_NO_SEQUENTIEL).ForeColor

  itmProjSoum.ListSubItems(I_COL_SOUM_NO_SEQUENTIEL).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_NO_SEQUENTIEL).Bold

  itmProjSoum.SubItems(I_COL_SOUM_PROVENANCE) = lvwSoumission.ListItems(iIndexCopie).SubItems(I_COL_SOUM_PROVENANCE)

  itmProjSoum.ListSubItems(I_COL_SOUM_PROVENANCE).ForeColor = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_PROVENANCE).ForeColor

  itmProjSoum.ListSubItems(I_COL_SOUM_PROVENANCE).Bold = lvwSoumission.ListItems(iIndexCopie).ListSubItems(I_COL_SOUM_PROVENANCE).Bold
  End If
  End If

  Call lvwSoumission.ListItems.Remove(iIndexCopie)

  Call lvwSoumission.Refresh

  iIndex = iIndex + 1
  End If

   iCompteur2 = iCompteur2 + 1
   Loop
7  Next iCompteur

7  Set rstSection = Nothing

7  If lvwSoumission.ListItems.count > 0 Then
7  Call Deselect

7  lvwSoumission.ListItems(1).Selected = True
7  End If

80 Exit Sub

Oups:

80 wOups "frmProjSoumElec", "UpdateOrdre", Err, Err.number, Err.Description
End Sub

Private Sub UpdatePieces()

 On Error GoTo Oups

 Dim rstPieceFRS As ADODB.Recordset
 Dim rstConfig As ADODB.Recordset
 Dim itmPiece As ListItem
 Dim iCompteur As Integer
 Dim sTauxUSA As String
 Dim sTauxSPA As String
 
 Set rstPieceFRS = New ADODB.Recordset
 
 Set rstConfig = New ADODB.Recordset

 Call rstConfig.Open("SELECT TauxAmericain, TauxEspagnol FROM GrbConfig", g_connData, adOpenForwardOnly, adLockReadOnly)

 sTauxUSA = rstConfig.Fields("TauxAmericain")
  sTauxSPA = rstConfig.Fields("TauxEspagnol")

  Call rstConfig.Close
  Set rstConfig = Nothing
 
  For iCompteur = 1 To lvwSoumission.ListItems.count
 'Si ce n'est pas une section
  If lvwSoumission.ListItems(iCompteur).Tag <> vbNullString Then
 'Si ce n'est pas une sous-section
  If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> vbNullString Then
  Set itmPiece = lvwSoumission.ListItems(iCompteur)
 
  Call ValeurParDefaut(itmPiece)
 
 If itmPiece.SubItems(I_COL_SOUM_PIECE) <> "Texte" And itmPiece.SubItems(I_COL_SOUM_PIECE) <> "Text" Then
 Call rstPieceFRS.Open("SELECT PRIX_LIST, PRIX_SP, PRIX_NET, ESCOMPTE, DeviseMonétaire FROM GrbPiecesFRS WHERE PIECE = '" & Replace(itmPiece.SubItems(I_COL_SOUM_PIECE), "'", "''") & "' AND IDFRS = " & itmPiece.ListSubItems(I_COL_SOUM_DISTRIB).Tag, g_connData, adOpenDynamic, adLockOptimistic)
 
 If Not rstPieceFRS.EOF Then
 If Not IsNull(rstPieceFRS.Fields("PRIX_LIST")) Then
 If Trim(rstPieceFRS.Fields("PRIX_LIST")) <> "" Then
 If rstPieceFRS.Fields("DeviseMonétaire") = "USA" Then
 itmPiece.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(rstPieceFRS.Fields("PRIX_LIST")) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
 Else
 If rstPieceFRS.Fields("DeviseMonétaire") = "SPA" Then
 itmPiece.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(rstPieceFRS.Fields("PRIX_LIST")) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
 Else
 itmPiece.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(rstPieceFRS.Fields("PRIX_LIST"), MODE_ARGENT, 4)
 End If
 End If
 Else
 itmPiece.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion("0", MODE_ARGENT, 4)
 End If
 Else
 itmPiece.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion("0", MODE_ARGENT, 4)
1  End If
 
 
 If Not IsNull(rstPieceFRS.Fields("PRIX_NET")) Then
 If Trim(rstPieceFRS.Fields("PRIX_NET")) <> vbNullString Then
 If rstPieceFRS.Fields("ESCOMPTE") <> vbNullString Then
 itmPiece.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion(rstPieceFRS.Fields("Escompte"), MODE_POURCENT)
 End If
 
 itmPiece.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(rstPieceFRS.Fields("PRIX_NET"), MODE_ARGENT, 4)
 Else
 If Not IsNull(rstPieceFRS.Fields("PRIX_SP")) Then
 itmPiece.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(rstPieceFRS.Fields("PRIX_SP"), MODE_ARGENT, 4)
 Else
 itmPiece.SubItems(I_COL_SOUM_PRIX_NET) = ""
 End If
 End If
 Else
 If Not IsNull(rstPieceFRS.Fields("PRIX_SP")) Then
 itmPiece.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(rstPieceFRS.Fields("PRIX_SP"), MODE_ARGENT, 4)
 Else
 itmPiece.SubItems(I_COL_SOUM_PRIX_NET) = ""
 End If
 End If

 
 If rstPieceFRS.Fields("DeviseMonétaire") = "USA" Then
 itmPiece.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(itmPiece.SubItems(I_COL_SOUM_PRIX_NET)) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
 Else
 If rstPieceFRS.Fields("DeviseMonétaire") = "SPA" Then
 itmPiece.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(itmPiece.SubItems(I_COL_SOUM_PRIX_NET)) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
 Else
 itmPiece.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(itmPiece.SubItems(I_COL_SOUM_PRIX_NET), MODE_ARGENT, 4)
 End If
 End If
 
 'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
 itmPiece.SubItems(I_COL_SOUM_TOTAL) = Conversion(CStr(Replace(itmPiece.Text, "*", vbNullString) * itmPiece.SubItems(I_COL_SOUM_PRIX_NET) * CSng(m_sProfit)), MODE_ARGENT)
 
 'Pour le profit, c'est le prix total - (prix net * quantité)
 itmPiece.SubItems(I_COL_SOUM_PROFIT) = Conversion(CStr(itmPiece.SubItems(I_COL_SOUM_TOTAL) - (itmPiece.SubItems(I_COL_SOUM_PRIX_NET) * Replace(itmPiece.Text, "*", vbNullString))), MODE_ARGENT)
 
 'Pour garder en mémoire le prix d'origine, je le mets dans le
 'tag de la colonne Prix Listé
 If Trim$(itmPiece.SubItems(I_COL_SOUM_PRIX_LIST)) = vbNullString Then
 itmPiece.SubItems(I_COL_SOUM_PRIX_LIST) = " "
 End If
 
 If Not IsNull(rstPieceFRS.Fields("PRIX_NET")) Then
 If Trim(rstPieceFRS.Fields("PRIX_NET")) <> vbNullString Then
 itmPiece.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = rstPieceFRS.Fields("PRIX_LIST")
 Else
 itmPiece.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = rstPieceFRS.Fields("PRIX_SP")
 End If
 Else
4 itmPiece.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = rstPieceFRS.Fields("PRIX_SP")
4 End If
4 Else
4 Call MsgBox("Il n'y a pas de prix du fournisseur " & itmPiece.SubItems(I_COL_SOUM_DISTRIB) & " pour la pièce " & itmPiece.SubItems(I_COL_SOUM_PIECE) & " ou la pièce n'existe plus!", vbOKOnly, "Erreur")
4 End If

4 Call rstPieceFRS.Close
4 End If
4 End If
4 End If
4 Next

4 Set rstPieceFRS = Nothing

4  Exit Sub

Oups:

4  wOups "frmProjSoumElec", "UpdatePieces", Err, Err.number, Err.Description
End Sub

Private Sub cmdCreerProjet_Click()

 On Error GoTo Oups

 'Créé un projet à partir d'une soumission
 Dim rstProjSoum As ADODB.Recordset
 Dim sNoProjet As String
 Dim sUser As String
 Dim iCompteur As Integer
 Dim bExiste As Boolean
 Dim bNoValide As Boolean
 Dim sLiaison As String

 If cmbProjSoum.ListCount > 0 Then
 If Right$(txtNoProjSoum.Text, 2) = "99" Then
 Call MsgBox("Impossible de créer un projet à partir de cette soumission!", vbOKOnly, "Erreur")
 
  Exit Sub
  End If

  Set rstProjSoum = New ADODB.Recordset

  Call rstProjSoum.Open("SELECT * FROM GrbProjSoum WHERE IDProjSoum = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

  If rstProjSoum.Fields("Ouvert") = False Or rstProjSoum.Fields("Verrouillé") = True Then
  If rstProjSoum.Fields("Ouvert") = False Then
  Call MsgBox("Cette soumission est fermée!", vbOKOnly, "Erreur")
  Else
 Call MsgBox("Cette soumission est verrouillée!", vbOKOnly, "Erreur")
End If
 
 Call rstProjSoum.Close
 Set rstProjSoum = Nothing

 Exit Sub
 End If

 Call rstProjSoum.Close

 If VerifierSiOuvert(sUser) = False Then
 'Demande du numéro de projet
 sNoProjet = InputBox("Quel est le numéro du projet?")

 If Trim$(sNoProjet) <> vbNullString Then
 Screen.MousePointer = vbHourglass

 bNoValide = True

 If ValiderFormatNumeroProjSoum(sNoProjet) = False Then
 bNoValide = False
 End If

 If bNoValide = True Then
 If ValiderFormatElectrique(sNoProjet) = False Then
 bNoValide = False
 End If
1  End If

 If bNoValide = True Then
 If ValiderFormatJobAvecSoum(sNoProjet) = False Then
 bNoValide = False
 End If
 End If

 If bNoValide = False Then
 Set rstProjSoum = Nothing

 Screen.MousePointer = vbDefault

 Exit Sub
 End If

 sNoProjet = UCase(sNoProjet)

 Call rstProjSoum.Open("SELECT * FROM GrbSoumissionElec WHERE IDSoumission = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 If rstProjSoum.EOF Then
 bExiste = False
 Else
 bExiste = True

 Call MsgBox("Ce numéro existe dans les soumissions électriques!", vbOKOnly, "Erreur")
 End If

 Call rstProjSoum.Close

 If bExiste = False Then
 Call rstProjSoum.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 If rstProjSoum.EOF Then
 bExiste = False
 Else
 bExiste = True

 Call MsgBox("Ce numéro existe dans les projets électriques!", vbOKOnly, "Erreur")
 End If

 Call rstProjSoum.Close
 End If

 If bExiste = False Then
 Call rstProjSoum.Open("SELECT * FROM GrbSoumissionMec WHERE IDSoumission = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)

 If rstProjSoum.EOF Then
 bExiste = False
 Else
 bExiste = True

 Call MsgBox("Ce numéro existe dans les soumissions mécaniques!", vbOKOnly, "Erreur")
 End If

 Call rstProjSoum.Close
 End If
 
 If bExiste = False Then
 Call rstProjSoum.Open("SELECT * FROM GrbProjetMec WHERE IDProjet = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)

4 If rstProjSoum.EOF Then
4 bExiste = False
4 Else
4 bExiste = True
 
4 Call MsgBox("Ce numéro existe dans les projets mécaniques!", vbOKOnly, "Erreur")
4 End If

4 Call rstProjSoum.Close
4 End If

4 If bExiste = True Then
4 Set rstProjSoum = Nothing

4 Screen.MousePointer = vbDefault

4  Exit Sub
4  End If

4  Call rstProjSoum.Open("SELECT * FROM GrbProjSoum WHERE IDProjSoum = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)

4  If Not rstProjSoum.EOF Then
4  If rstProjSoum.Fields("Ouvert") = False Then
4  Call MsgBox("Ce numéro est fermé!", vbOKOnly, "Erreur")

4  Call rstProjSoum.Close
4  Set rstProjSoum = Nothing

50 Exit Sub
 End If
 End If

 Call rstProjSoum.Close
 Set rstProjSoum = Nothing
 
 If Right$(sNoProjet, 2) >= 60 And Right$(sNoProjet, 2) <=   Then
 sLiaison = InputBox("Quelle est l'extention du projet " & Left$(sNoProjet, 6) & " auquel ce projet sera lié?")
 End If
 
 Call frmChoixTransfertJob.Afficher(txtNoProjSoum.Text, "E")
 
 If m_bTransfertJobCancel = False Then
 'Appel de la méthode pour créer le projet
 Call TransfererSoumDansProjet(sNoProjet, sLiaison)
 
 'On affiche le projet qui vient d'être créé
 If m_bComboChoix = True Then
5  cmbChoix.ListIndex = I_IDX_PROJET
 
5  For iCompteur = 0 To cmbProjSoum.ListCount - 1
5  If cmbProjSoum.LIST(iCompteur) = sNoProjet Then
5  cmbProjSoum.ListIndex = iCompteur
 
5  Exit For
5  End If
5  Next

5  If sLiaison <> "" Then
60 For iCompteur = 1 To lvwSoumission.ListItems.count
 'Si pas une section
  If lvwSoumission.ListItems(iCompteur).Tag <> "" Then
 'Si pas une sous-section
  If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> "" Then
  If Right$(sNoProjet, 2) >= 60 And Right$(sNoProjet, 2) <= 7 Then
  Call AjouterPiecesExtraChargeableDansJob(lvwSoumission.ListItems(iCompteur), sLiaison)
  Else
  If Right$(sNoProjet, 2) >= 80 And Right$(sNoProjet, 2) <=   Then
  Call AjouterPiecesExtraDansJob(lvwSoumission.ListItems(iCompteur), sLiaison)
  End If
  End If
  End If
  End If

6  Call CalculerTotalRecordset(sNoProjet)
6  Next
6  End If

6  Call AjouterProjetAuCumulatif
6  End If
 
 'Il faut enlever le bouton puisqu'on ne peut pas créer plus
 'd'un projet avec une soumission
6  cmdCreerProjet.Visible = False
6  End If
 
6  Screen.MousePointer = vbDefault
70 Else
  Set rstProjSoum = Nothing

  Exit Sub
  End If
  Else
  Call MsgBox("Cette soumission est en modification par " & sUser & "!", vbOKOnly, "Erreur")
  End If
  End If

  Exit Sub

Oups:

  wOups "frmProjSoumElec", "cmdCreerProjet_Click", Err, Err.number, Err.Description
End Sub

Private Function VerifierSiDejaProjet() As Boolean

 On Error GoTo Oups

 'Méthode qui sert à vérifier si une soumission est déjà assignée à un projet
 Dim rstProjet As ADODB.Recordset
 
 Set rstProjet = New ADODB.Recordset
 
 Call rstProjet.Open("SELECT * FROM GrbProjetElec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 If Not rstProjet.EOF Then
 VerifierSiDejaProjet = True
 End If
 
 Call rstProjet.Close
 Set rstProjet = Nothing

 Exit Function

Oups:

 wOups "frmProjSoumElec", "VerifierSiDejaProjet", Err, Err.number, Err.Description
End Function

Private Sub TransfererSoumDansProjet(ByVal sNoProjet As String, ByVal sLiaison As String)

 On Error GoTo Oups

 'Méthode qui transfère les données de la soumission dans les tables
 'GrbProjet et Grbpièces
 Dim rstSoum As ADODB.Recordset
 Dim rstProjet As ADODB.Recordset
 Dim rstSoumPiece As ADODB.Recordset
 Dim rstProjetPiece As ADODB.Recordset
 Dim rstEmploye As ADODB.Recordset
 Dim rstProjSoum As ADODB.Recordset
 Dim rstConfig As ADODB.Recordset
 Dim iCompteur As Integer

 Set rstSoum = New ADODB.Recordset
 Set rstProjet = New ADODB.Recordset
  Set rstSoumPiece = New ADODB.Recordset
  Set rstProjetPiece = New ADODB.Recordset
  Set rstEmploye = New ADODB.Recordset
  Set rstProjSoum = New ADODB.Recordset

 'Ouverture de la soumission
  Call rstSoum.Open("SELECT * FROM GrbSoumissionElec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
  Call rstSoumPiece.Open("SELECT * FROM GrbSoumission_Pieces WHERE IDSoumission = '" & txtNoProjSoum.Text & "' AND Type = 'E' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Ouverture du projet
  Call rstProjet.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)
  Call rstProjetPiece.Open("SELECT * FROM GrbProjet_Pieces", g_connData, adOpenDynamic, adLockOptimistic)

10 Call rstProjSoum.Open("SELECT * FROM GrbProjSoum WHERE IDProjSoum = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
If rstProjSoum.EOF Then
 Call rstProjSoum.AddNew
 
 rstProjSoum.Fields("IDProjSoum") = sNoProjet
 rstProjSoum.Fields("NoClient") = rstSoum.Fields("IDClient")
 rstProjSoum.Fields("DateOuverture") = ConvertDate(Date)
 rstProjSoum.Fields("Description") = rstSoum.Fields("Description")
 rstProjSoum.Fields("Ouvert") = True
 rstProjSoum.Fields("Type") = "P"
 
 Call rstProjSoum.Update
End If
 
Call rstProjSoum.Close

1  Call rstProjSoum.Open("SELECT * FROM GrbProjSoum WHERE IDProjSoum = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

rstProjSoum.Fields("Ouvert") = False

 Call rstProjSoum.Update

Call rstProjSoum.Close
 Set rstProjSoum = Nothing
 
 'On l'ajoute
Call rstProjet.AddNew
 
 rstProjet.Fields("IDProjet") = sNoProjet
1  rstProjet.Fields("IDSoumission") = rstSoum.Fields("IDSoumission")
 rstProjet.Fields("IDClient") = rstSoum.Fields("IDClient")
 rstProjet.Fields("IDContact") = rstSoum.Fields("IDContact")
rstProjet.Fields("Description") = rstSoum.Fields("Description")
rstProjet.Fields("Panneau_aire") = rstSoum.Fields("Panneau_aire")
rstProjet.Fields("panneau_espace") = rstSoum.Fields("panneau_espace")
rstProjet.Fields("NbreManuel") = rstSoum.Fields("NbreManuel")
rstProjet.Fields("transport") = rstSoum.Fields("transport")
rstProjet.Fields("csa") = rstSoum.Fields("csa")
rstProjet.Fields("cul") = rstSoum.Fields("cul")
rstProjet.Fields("cur") = rstSoum.Fields("cur")
rstProjet.Fields("ul") = rstSoum.Fields("ul")
rstProjet.Fields("ur") = rstSoum.Fields("ur")
2  rstProjet.Fields("ce") = rstSoum.Fields("ce")
rstProjet.Fields("Delais") = rstSoum.Fields("Delais")
2  rstProjet.Fields("Creer") = ConvertDate(Date)
rstProjet.Fields("CheminPhotos") = rstSoum.Fields("CheminPhotos")

2  If sLiaison <> "" Then
 rstProjet.Fields("LiaisonChargeable") = sLiaison
2  End If

Call rstEmploye.Open("SELECT noEmploye FROM GrbEmployés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
30 rstProjet.Fields("Creer_Par") = rstEmploye.Fields("noEmploye")
 
Call rstEmploye.Close
Set rstEmploye = Nothing
 
rstProjet.Fields("TempsDessin") = rstSoum.Fields("TempsDessin")
rstProjet.Fields("TempsFabrication") = rstSoum.Fields("TempsFabrication")
rstProjet.Fields("TempsAssemblage") = rstSoum.Fields("TempsAssemblage")
rstProjet.Fields("TempsProgInterface") = rstSoum.Fields("TempsProgInterface")
rstProjet.Fields("TempsProgAutomate") = rstSoum.Fields("TempsProgAutomate")
rstProjet.Fields("TempsProgRobot") = rstSoum.Fields("TempsProgRobot")
rstProjet.Fields("TempsVision") = rstSoum.Fields("TempsVision")
rstProjet.Fields("TempsTest") = rstSoum.Fields("TempsTest")
rstProjet.Fields("TempsInstallation") = 0
3  rstProjet.Fields("TempsMiseService") = 0
rstProjet.Fields("TempsFormation") = rstSoum.Fields("TempsFormation")
3  rstProjet.Fields("TempsGestion") = rstSoum.Fields("TempsGestion")
rstProjet.Fields("TempsShipping") = rstSoum.Fields("TempsShipping")

3  Set rstConfig = New ADODB.Recordset

If Not IsNull(rstSoum.Fields("TauxDessin")) Then
 rstProjet.Fields("TauxDessin") = rstSoum.Fields("TauxDessin")
 Else
Call rstConfig.Open("SELECT TauxDessinElec FROM GrbConfig", g_connData, adOpenForwardOnly, adLockReadOnly)

4 rstProjet.Fields("TauxDessin") = rstConfig.Fields("TauxDessinElec")

4 Call rstConfig.Close
4 End If

4 If Not IsNull(rstSoum.Fields("TauxFabrication")) Then
4 rstProjet.Fields("TauxFabrication") = rstSoum.Fields("TauxFabrication")
4 Else
4 Call rstConfig.Open("SELECT TauxFabrication FROM GrbConfig", g_connData, adOpenForwardOnly, adLockReadOnly)

4 rstProjet.Fields("TauxFabrication") = rstConfig.Fields("TauxFabrication")

4 Call rstConfig.Close
4 End If

4 If Not IsNull(rstSoum.Fields("TauxAssemblage")) Then
4  rstProjet.Fields("TauxAssemblage") = rstSoum.Fields("TauxAssemblage")
4  Else
4  Call rstConfig.Open("SELECT TauxAssemblageElec FROM GrbConfig", g_connData, adOpenForwardOnly, adLockReadOnly)

4  rstProjet.Fields("TauxAssemblage") = rstConfig.Fields("TauxAssemblageElec")

4  Call rstConfig.Close
4  End If

4  If Not IsNull(rstSoum.Fields("TauxProgInterface")) Then
4  rstProjet.Fields("TauxProgInterface") = rstSoum.Fields("TauxProgInterface")
50 Else
5 Call rstConfig.Open("SELECT TauxProgInterface FROM GrbConfig", g_connData, adOpenForwardOnly, adLockReadOnly)

 rstProjet.Fields("TauxProgInterface") = rstConfig.Fields("TauxProgInterface")

 Call rstConfig.Close
 End If

 If Not IsNull(rstSoum.Fields("TauxProgAutomate")) Then
 rstProjet.Fields("TauxProgAutomate") = rstSoum.Fields("TauxProgAutomate")
 Else
 Call rstConfig.Open("SELECT TauxProgAutomate FROM GrbConfig", g_connData, adOpenForwardOnly, adLockReadOnly)

 rstProjet.Fields("TauxProgAutomate") = rstConfig.Fields("TauxProgAutomate")

 Call rstConfig.Close
 End If

5  If Not IsNull(rstSoum.Fields("TauxProgRobot")) Then
5  rstProjet.Fields("TauxProgRobot") = rstSoum.Fields("TauxProgRobot")
5  Else
5  Call rstConfig.Open("SELECT TauxProgRobot FROM GrbConfig", g_connData, adOpenForwardOnly, adLockReadOnly)

5  rstProjet.Fields("TauxProgRobot") = rstConfig.Fields("TauxProgRobot")

5  Call rstConfig.Close
5  End If

5  If Not IsNull(rstSoum.Fields("TauxVision")) Then
60 rstProjet.Fields("TauxVision") = rstSoum.Fields("TauxVision")
60 Else
  Call rstConfig.Open("SELECT TauxVision FROM GrbConfig", g_connData, adOpenForwardOnly, adLockReadOnly)

  rstProjet.Fields("TauxVision") = rstConfig.Fields("TauxVision")

  Call rstConfig.Close
  End If

  If Not IsNull(rstSoum.Fields("TauxTest")) Then
  rstProjet.Fields("TauxTest") = rstSoum.Fields("TauxTest")
  Else
  Call rstConfig.Open("SELECT TauxTestElec FROM GrbConfig", g_connData, adOpenForwardOnly, adLockReadOnly)

  rstProjet.Fields("TauxTest") = rstConfig.Fields("TauxTestElec")

  Call rstConfig.Close
6  End If

6  If Not IsNull(rstSoum.Fields("TauxInstallation")) Then
6  rstProjet.Fields("TauxInstallation") = rstSoum.Fields("TauxInstallation")
6  Else
6  Call rstConfig.Open("SELECT TauxInstallationElec FROM GrbConfig", g_connData, adOpenForwardOnly, adLockReadOnly)

6  rstProjet.Fields("TauxInstallation") = rstConfig.Fields("TauxInstallationElec")

6  Call rstConfig.Close
6  End If

70 If Not IsNull(rstSoum.Fields("TauxMiseService")) Then
  rstProjet.Fields("TauxMiseService") = rstSoum.Fields("TauxMiseService")
  Else
  Call rstConfig.Open("SELECT TauxMiseService FROM GrbConfig", g_connData, adOpenForwardOnly, adLockReadOnly)

  rstProjet.Fields("TauxMiseService") = rstConfig.Fields("TauxMiseService")

  Call rstConfig.Close
  End If

  If Not IsNull(rstSoum.Fields("TauxFormation")) Then
  rstProjet.Fields("TauxFormation") = rstSoum.Fields("TauxFormation")
  Else
  Call rstConfig.Open("SELECT TauxFormationElec FROM GrbConfig", g_connData, adOpenForwardOnly, adLockReadOnly)

  rstProjet.Fields("TauxFormation") = rstConfig.Fields("TauxFormationElec")

   Call rstConfig.Close
   End If

7  If Not IsNull(rstSoum.Fields("TauxGestion")) Then
7  rstProjet.Fields("TauxGestion") = rstSoum.Fields("TauxGestion")
7  Else
7  Call rstConfig.Open("SELECT TauxGestionProjetsElec FROM GrbConfig", g_connData, adOpenForwardOnly, adLockReadOnly)

7  rstProjet.Fields("TauxGestion") = rstConfig.Fields("TauxGestionProjetsElec")

7  Call rstConfig.Close
80 End If

80 If Not IsNull(rstSoum.Fields("TauxShipping")) Then
  rstProjet.Fields("TauxShipping") = rstSoum.Fields("TauxShipping")
  Else
  Call rstConfig.Open("SELECT TauxShippingElec FROM GrbConfig", g_connData, adOpenForwardOnly, adLockReadOnly)

  rstProjet.Fields("TauxShipping") = rstConfig.Fields("TauxShippingElec")

  Call rstConfig.Close
  End If

  Set rstConfig = Nothing

  rstProjet.Fields("PrixEmballage") = rstSoum.Fields("PrixEmballage")

  rstProjet.Fields("imprevue") = rstSoum.Fields("imprevue")
  rstProjet.Fields("commission") = rstSoum.Fields("commission")
   rstProjet.Fields("Profit") = rstSoum.Fields("Profit")
   rstProjet.Fields("total_manuel") = rstSoum.Fields("total_manuel")
   rstProjet.Fields("total_commission") = rstSoum.Fields("total_Commission")
   rstProjet.Fields("total_profit") = rstSoum.Fields("Total_Profit")
8  rstProjet.Fields("PrixTotal") = rstSoum.Fields("PrixTotal")
8  rstProjet.Fields("total_piece") = rstSoum.Fields("Total_piece")
8  rstProjet.Fields("total_imprevue") = rstSoum.Fields("total_imprevue")
8  rstProjet.Fields("total_temps") = rstSoum.Fields("total_temps")
90 rstProjet.Fields("SansTemps") = rstSoum.Fields("SansTemps")
90 rstProjet.Fields("MontantForfait") = rstSoum.Fields("MontantForfait")
  rstProjet.Fields("InitialeForfait") = rstSoum.Fields("InitialeForfait")
  rstProjet.Fields("ProchaineCommande") = 1

  Call rstProjet.Update
 
 'Ajout des pièces
  Do While Not rstSoumPiece.EOF
  If rstSoumPiece.Fields("TransfertJob") = True Then
  Call rstProjetPiece.AddNew

  rstProjetPiece.Fields("Type") = "E"
 
  rstProjetPiece.Fields("IDProjet") = sNoProjet
  rstProjetPiece.Fields("IDSection") = rstSoumPiece.Fields("IDSection")
  rstProjetPiece.Fields("NumItem") = rstSoumPiece.Fields("NumItem")
 rstProjetPiece.Fields("Qté") = rstSoumPiece.Fields("Qté")
   rstProjetPiece.Fields("Desc_FR") = rstSoumPiece.Fields("Desc_FR")
 rstProjetPiece.Fields("Desc_EN") = rstSoumPiece.Fields("Desc_EN")
   rstProjetPiece.Fields("Manufact") = rstSoumPiece.Fields("Manufact")
 rstProjetPiece.Fields("Prix_List") = rstSoumPiece.Fields("Prix_list")
   rstProjetPiece.Fields("Escompte") = rstSoumPiece.Fields("Escompte")
 rstProjetPiece.Fields("Prix_net") = rstSoumPiece.Fields("Prix_net")
9  rstProjetPiece.Fields("OrdreSection") = rstSoumPiece.Fields("OrdreSection")
 rstProjetPiece.Fields("NuméroLigne") = rstSoumPiece.Fields("NuméroLigne")
10 rstProjetPiece.Fields("IDFRS") = rstSoumPiece.Fields("IDFRS")
1 rstProjetPiece.Fields("Temps") = rstSoumPiece.Fields("Temps")
1 rstProjetPiece.Fields("Temps_total") = rstSoumPiece.Fields("Temps_Total")
 rstProjetPiece.Fields("Prix_total") = rstSoumPiece.Fields("Prix_Total")
1 rstProjetPiece.Fields("Profit_argent") = rstSoumPiece.Fields("Profit_argent")
 rstProjetPiece.Fields("SousSection") = rstSoumPiece.Fields("SousSection")
1 rstProjetPiece.Fields("PrixOrigine") = rstSoumPiece.Fields("PrixOrigine")
 rstProjetPiece.Fields("Visible") = rstSoumPiece.Fields("Visible")
1 rstProjetPiece.Fields("Commentaire") = rstSoumPiece.Fields("Commentaire")
 rstProjetPiece.Fields("Quoté") = rstSoumPiece.Fields("Quoté")
1 rstProjetPiece.Fields("Devise") = rstSoumPiece.Fields("Devise")

10  Call rstProjetPiece.Update

10  If sLiaison <> "" Then
10  If Right$(sNoProjet, 2) >= "60" And Right$(sNoProjet, 2) <= 7 Then
1075
10  Else
10  If Right$(sNoProjet, 2) >= 80 And Right$(sNoProjet, 2) <=   Then
1090
10  End If
1 End If
11 End If
1 End If
 
1 Call rstSoumPiece.MoveNext
11 Loop

11 m_eType = TYPE_PROJET

11 If CDbl(rstSoum.Fields("TempsInstallation")) > 0 Or CDbl(rstSoum.Fields("TempsMiseService")) > 0 Then
1 Call CreerProjetInstallation(Left$(sNoProjet, 7) & "51")
11 End If

11 Call rstSoum.Close
11 Set rstSoum = Nothing

11 Call rstProjet.Close
11  Set rstProjet = Nothing

11  Call rstSoumPiece.Close
1 Set rstSoumPiece = Nothing
 
11  Call rstProjetPiece.Close
1 Set rstProjetPiece = Nothing

11  Call CalculerTotalRecordset(sNoProjet)

1 Exit Sub

Oups:

11  wOups "frmProjSoumElec", "TransfererSoumDansProjet", Err, Err.number, Err.Description
End Sub

Private Sub CreerProjetInstallation(ByVal sNoProjet As String)

 On Error GoTo Oups

 Dim rstSoum As ADODB.Recordset
 Dim rstProjet As ADODB.Recordset
 Dim rstEmploye As ADODB.Recordset
 Dim rstProjSoum As ADODB.Recordset
 Dim rstConfig As ADODB.Recordset
 Dim iCompteur As Integer

 Set rstSoum = New ADODB.Recordset
 Set rstProjet = New ADODB.Recordset
 Set rstEmploye = New ADODB.Recordset
 Set rstProjSoum = New ADODB.Recordset

 'Ouverture de la soumission
  Call rstSoum.Open("SELECT * FROM GrbSoumissionElec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Ouverture du projet
  Call rstProjet.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)

  If rstProjet.EOF Then
  Call rstProjSoum.Open("SELECT * FROM GrbProjSoum WHERE IDProjSoum = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
  If rstProjSoum.EOF Then
  Call rstProjSoum.AddNew
 
  rstProjSoum.Fields("IDProjSoum") = sNoProjet
  rstProjSoum.Fields("NoClient") = rstSoum.Fields("IDClient")
 rstProjSoum.Fields("Description") = rstSoum.Fields("Description")
rstProjSoum.Fields("DateOuverture") = ConvertDate(Date)
 rstProjSoum.Fields("Ouvert") = True
 rstProjSoum.Fields("Type") = "P"
 
 Call rstProjSoum.Update
 End If
 
 Call rstProjSoum.Close
 Set rstProjSoum = Nothing
 
 'On l'ajoute
 Call rstProjet.AddNew
 
 rstProjet.Fields("IDProjet") = sNoProjet
 rstProjet.Fields("IDSoumission") = vbNullString
 rstProjet.Fields("IDClient") = rstSoum.Fields("IDClient")
rstProjet.Fields("IDContact") = rstSoum.Fields("IDContact")
 rstProjet.Fields("Description") = rstSoum.Fields("Description")
 rstProjet.Fields("Panneau_aire") = rstSoum.Fields("Panneau_aire")
 rstProjet.Fields("panneau_espace") = rstSoum.Fields("panneau_espace")
 rstProjet.Fields("NbreManuel") = rstSoum.Fields("NbreManuel")
 rstProjet.Fields("transport") = rstSoum.Fields("transport")
 rstProjet.Fields("csa") = rstSoum.Fields("csa")
1  rstProjet.Fields("cul") = rstSoum.Fields("cul")
 rstProjet.Fields("cur") = rstSoum.Fields("cur")
 rstProjet.Fields("ul") = rstSoum.Fields("ul")
 rstProjet.Fields("ur") = rstSoum.Fields("ur")
 rstProjet.Fields("ce") = rstSoum.Fields("ce")
 rstProjet.Fields("Delais") = rstSoum.Fields("Delais")
 rstProjet.Fields("Creer") = ConvertDate(Date)
 rstProjet.Fields("CheminPhotos") = rstSoum.Fields("CheminPhotos")
 
 Call rstEmploye.Open("SELECT noEmploye FROM GrbEmployés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 rstProjet.Fields("Creer_Par") = rstEmploye.Fields("noEmploye")
 
 Call rstEmploye.Close
 Set rstEmploye = Nothing
 
 rstProjet.Fields("TempsDessin") = 0
rstProjet.Fields("TempsFabrication") = 0
 rstProjet.Fields("TempsAssemblage") = 0
rstProjet.Fields("TempsProgInterface") = 0
 rstProjet.Fields("TempsProgAutomate") = 0
rstProjet.Fields("TempsProgRobot") = 0
 rstProjet.Fields("TempsVision") = 0
rstProjet.Fields("TempsTest") = 0
 rstProjet.Fields("TempsInstallation") = rstSoum.Fields("TempsInstallation")
rstProjet.Fields("TempsMiseService") = rstSoum.Fields("TempsMiseService")
3 rstProjet.Fields("TempsFormation") = 0
 rstProjet.Fields("TempsGestion") = 0
 rstProjet.Fields("TempsShipping") = 0

 Set rstConfig = New ADODB.Recordset

 If Not IsNull(rstSoum.Fields("TauxDessin")) Then
 rstProjet.Fields("TauxDessin") = rstSoum.Fields("TauxDessin")
 Else
 Call rstConfig.Open("SELECT TauxDessinElec FROM GrbConfig", g_connData, adOpenForwardOnly, adLockReadOnly)

 rstProjet.Fields("TauxDessin") = rstConfig.Fields("TauxDessinElec")

 Call rstConfig.Close
 End If

If Not IsNull(rstSoum.Fields("TauxFabrication")) Then
 rstProjet.Fields("TauxFabrication") = rstSoum.Fields("TauxFabrication")
Else
 Call rstConfig.Open("SELECT TauxFabrication FROM GrbConfig", g_connData, adOpenForwardOnly, adLockReadOnly)

 rstProjet.Fields("TauxFabrication") = rstConfig.Fields("TauxFabrication")

 Call rstConfig.Close
 End If

 If Not IsNull(rstSoum.Fields("TauxAssemblage")) Then
 rstProjet.Fields("TauxAssemblage") = rstSoum.Fields("TauxAssemblage")
4 Else
4 Call rstConfig.Open("SELECT TauxAssemblageElec FROM GrbConfig", g_connData, adOpenForwardOnly, adLockReadOnly)

4 rstProjet.Fields("TauxAssemblage") = rstConfig.Fields("TauxAssemblageElec")

4 Call rstConfig.Close
4 End If

4 If Not IsNull(rstSoum.Fields("TauxProgInterface")) Then
4 rstProjet.Fields("TauxProgInterface") = rstSoum.Fields("TauxProgInterface")
4 Else
4 Call rstConfig.Open("SELECT TauxProgInterface FROM GrbConfig", g_connData, adOpenForwardOnly, adLockReadOnly)

4 rstProjet.Fields("TauxProgInterface") = rstConfig.Fields("TauxProgInterface")

4 Call rstConfig.Close
4  End If

4  If Not IsNull(rstSoum.Fields("TauxProgAutomate")) Then
4  rstProjet.Fields("TauxProgAutomate") = rstSoum.Fields("TauxProgAutomate")
4  Else
4  Call rstConfig.Open("SELECT TauxProgAutomate FROM GrbConfig", g_connData, adOpenForwardOnly, adLockReadOnly)

4  rstProjet.Fields("TauxProgAutomate") = rstConfig.Fields("TauxProgAutomate")

4  Call rstConfig.Close
4  End If

50 If Not IsNull(rstSoum.Fields("TauxProgRobot")) Then
rstProjet.Fields("TauxProgRobot") = rstSoum.Fields("TauxProgRobot")
 Else
 Call rstConfig.Open("SELECT TauxProgRobot FROM GrbConfig", g_connData, adOpenForwardOnly, adLockReadOnly)

 rstProjet.Fields("TauxProgRobot") = rstConfig.Fields("TauxProgRobot")

 Call rstConfig.Close
 End If

 If Not IsNull(rstSoum.Fields("TauxVision")) Then
 rstProjet.Fields("TauxVision") = rstSoum.Fields("TauxVision")
 Else
 Call rstConfig.Open("SELECT TauxVision FROM GrbConfig", g_connData, adOpenForwardOnly, adLockReadOnly)

 rstProjet.Fields("TauxVision") = rstConfig.Fields("TauxVision")

5  Call rstConfig.Close
5  End If

5  If Not IsNull(rstSoum.Fields("TauxTest")) Then
5  rstProjet.Fields("TauxTest") = rstSoum.Fields("TauxTest")
5  Else
5  Call rstConfig.Open("SELECT TauxTestElec FROM GrbConfig", g_connData, adOpenForwardOnly, adLockReadOnly)

5  rstProjet.Fields("TauxTest") = rstConfig.Fields("TauxTestElec")

5  Call rstConfig.Close
60 End If

  If Not IsNull(rstSoum.Fields("TauxInstallation")) Then
  rstProjet.Fields("TauxInstallation") = rstSoum.Fields("TauxInstallation")
  Else
  Call rstConfig.Open("SELECT TauxInstallationElec FROM GrbConfig", g_connData, adOpenForwardOnly, adLockReadOnly)

  rstProjet.Fields("TauxInstallation") = rstConfig.Fields("TauxInstallationElec")

  Call rstConfig.Close
  End If

  If Not IsNull(rstSoum.Fields("TauxMiseService")) Then
  rstProjet.Fields("TauxMiseService") = rstSoum.Fields("TauxMiseService")
  Else
  Call rstConfig.Open("SELECT TauxMiseService FROM GrbConfig", g_connData, adOpenForwardOnly, adLockReadOnly)

6  rstProjet.Fields("TauxMiseService") = rstConfig.Fields("TauxMiseService")

6  Call rstConfig.Close
6  End If

6  If Not IsNull(rstSoum.Fields("TauxFormation")) Then
6  rstProjet.Fields("TauxFormation") = rstSoum.Fields("TauxFormation")
6  Else
6  Call rstConfig.Open("SELECT TauxFormationElec FROM GrbConfig", g_connData, adOpenForwardOnly, adLockReadOnly)

6  rstProjet.Fields("TauxFormation") = rstConfig.Fields("TauxFormationElec")

70 Call rstConfig.Close
  End If

  If Not IsNull(rstSoum.Fields("TauxGestion")) Then
  rstProjet.Fields("TauxGestion") = rstSoum.Fields("TauxGestion")
  Else
  Call rstConfig.Open("SELECT TauxGestionProjetsElec FROM GrbConfig", g_connData, adOpenForwardOnly, adLockReadOnly)

  rstProjet.Fields("TauxGestion") = rstConfig.Fields("TauxGestionProjetsElec")

  Call rstConfig.Close
  End If

  If Not IsNull(rstSoum.Fields("TauxShipping")) Then
  rstProjet.Fields("TauxShipping") = rstSoum.Fields("TauxShipping")
  Else
   Call rstConfig.Open("SELECT TauxShippingElec FROM GrbConfig", g_connData, adOpenForwardOnly, adLockReadOnly)

   rstProjet.Fields("TauxShipping") = rstConfig.Fields("TauxShippingElec")

7  Call rstConfig.Close
7  End If

7  Set rstConfig = Nothing

7  rstProjet.Fields("PrixEmballage") = 0

7  rstProjet.Fields("imprevue") = rstSoum.Fields("imprevue")
7  rstProjet.Fields("commission") = rstSoum.Fields("commission")
80 rstProjet.Fields("Profit") = rstSoum.Fields("Profit")
  rstProjet.Fields("total_manuel") = rstSoum.Fields("total_manuel")
  rstProjet.Fields("total_commission") = rstSoum.Fields("total_Commission")
  rstProjet.Fields("total_profit") = rstSoum.Fields("Total_Profit")
  rstProjet.Fields("PrixTotal") = rstSoum.Fields("PrixTotal")
  rstProjet.Fields("total_piece") = rstSoum.Fields("Total_piece")
  rstProjet.Fields("total_imprevue") = rstSoum.Fields("total_imprevue")
  rstProjet.Fields("total_temps") = rstSoum.Fields("total_temps")
  rstProjet.Fields("SansTemps") = rstSoum.Fields("SansTemps")
  rstProjet.Fields("ProchaineCommande") = 1

  Call rstProjet.Update
  End If
 
   Call rstSoum.Close
   Set rstSoum = Nothing

   Call rstProjet.Close
   Set rstProjet = Nothing

8  Call CalculerTotalRecordset(sNoProjet)

8  Exit Sub

Oups:

8  wOups "frmProjSoumElec", "CreerProjetInstallation", Err, Err.number, Err.Description
End Sub

Private Sub cmdDate_Click()

 On Error GoTo Oups

 'Ouverture du calendrier
 If Trim$(txtDelais.Text) <> vbNullString Then
 mvwDate.Value = txtDelais.Text
 Else
 mvwDate.Value = Date
 End If
 
 mvwDate.Visible = True
 
 Call mvwDate.SetFocus

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "cmdDate_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdDateFacturation_Click()

 On Error GoTo Oups

 'Ouverture du calendrier
 If txtDateFacturation.Text <> vbNullString Then
 mvwDateFacturation.Value = txtDateFacturation.Text
 Else
 mvwDateFacturation.Value = Date
 End If

 mvwDateFacturation.Visible = True

 Call mvwDateFacturation.SetFocus

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "cmdDateFacturation_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdDemande_Click()

 On Error GoTo Oups

 Dim rstProjSoum As ADODB.Recordset
 Dim sUser As String

 If Right$(txtNoProjSoum.Text, 2) = "99" Then
 Call MsgBox("Vous ne pouvez pas commander de pièce à partir de ce projet!", vbOKOnly, "Erreur")

 Exit Sub
 End If

 Set rstProjSoum = New ADODB.Recordset

 Call rstProjSoum.Open("SELECT * FROM GrbProjSoum WHERE IDProjSoum = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

 If rstProjSoum.Fields("Ouvert") = True And rstProjSoum.Fields("Verrouillé") = False Then
 If VerifierSiOuvert(sUser) = False Then
  Call frmChoixDemande.AfficherProjetSoumission(txtNoProjSoum.Text, ELECTRIQUE, MODE_PIECE, m_eType)
  Else
  If m_eType = TYPE_PROJET Then
  Call MsgBox("Ce projet est en modification par " & sUser & "!", vbOKOnly, "Erreur")
  Else
  Call MsgBox("Cette soumission est en modification par " & sUser & "!", vbOKOnly, "Erreur")
  End If
  End If
10 Else
1 If rstProjSoum.Fields("Ouvert") = False Then
 If m_eType = TYPE_PROJET Then
 Call MsgBox("Ce projet est fermé!", vbOKOnly, "Erreur")
 Else
 Call MsgBox("Cette soumission est fermée!", vbOKOnly, "Erreur")
 End If
 Else
 If m_eType = TYPE_PROJET Then
 Call MsgBox("Ce projet est verrouillé!", vbOKOnly, "Erreur")
 Else
 Call MsgBox("Cette soumission est verrouillée!", vbOKOnly, "Erreur")
 End If
 End If
 End If

Call rstProjSoum.Close
 Set rstProjSoum = Nothing
 
Exit Sub

Oups:

 wOups "frmProjSoumElec", "cmdDemande_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdExtra_Click()

 On Error GoTo Oups

 Dim sNumero As String
 Dim rstProjSoum As ADODB.Recordset
 Dim bExiste As Boolean
 Dim sExtension As String
 Dim bNoValide As Boolean

 If Right$(txtNoProjSoum.Text, 2) = "99" Then
 Call MsgBox("Vous ne pouvez pas faire un extra à partir de ce projet!", vbOKOnly, "Erreur")

 Exit Sub
 End If

 sExtension = Right$(txtNoProjSoum.Text, 2)
 
  sNumero = InputBox("Quel est l'extension à ajouter au numéro " & Left$(txtNoProjSoum.Text, 6) & "?")
 
  If sNumero <> vbNullString Then
  If Not IsNumeric(sNumero) Then
  Call MsgBox("Numéro non numérique!", vbOKOnly, "Erreur")
 
  Exit Sub
  End If
 
  sNumero = Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 3) & "-" & sNumero
 
  Screen.MousePointer = vbHourglass

bNoValide = True

1 If ValiderFormatNumeroProjSoum(sNumero) = False Then
 bNoValide = False
 End If

 If bNoValide = True Then
 If ValiderFormatElectrique(sNumero) = False Then
 bNoValide = False
 End If
 End If

 If bNoValide = True Then
 If ValiderFormatJobExtra(sNumero) = False Then
 bNoValide = False
 End If
 End If

 If bNoValide = False Then
 Screen.MousePointer = vbDefault

 Exit Sub
 End If

 sNumero = UCase(sNumero)

1  Set rstProjSoum = New ADODB.Recordset
 
 Call rstProjSoum.Open("SELECT * FROM GrbSoumissionElec WHERE IDSoumission = '" & sNumero & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 If rstProjSoum.EOF Then
 bExiste = False
 Else
 bExiste = True

 Call MsgBox("Ce numéro existe dans les soumissions électriques!", vbOKOnly, "Erreur")
 End If

 Call rstProjSoum.Close

 If bExiste = False Then
 Call rstProjSoum.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = '" & sNumero & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 If rstProjSoum.EOF Then
 bExiste = False
 Else
 bExiste = True

 Call MsgBox("Ce numéro existe dans les projets électriques!", vbOKOnly, "Erreur")
 End If

 Call rstProjSoum.Close
 End If

If bExiste = False Then
 Call rstProjSoum.Open("SELECT * FROM GrbSoumissionMec WHERE IDSoumission = '" & sNumero & "'", g_connData, adOpenDynamic, adLockOptimistic)

 If rstProjSoum.EOF Then
 bExiste = False
 Else
 bExiste = True

 Call MsgBox("Ce numéro existe dans les soumissions mécaniques!", vbOKOnly, "Erreur")
 End If

 Call rstProjSoum.Close
 End If
 
 If bExiste = False Then
 Call rstProjSoum.Open("SELECT * FROM GrbProjetMec WHERE IDProjet = '" & sNumero & "'", g_connData, adOpenDynamic, adLockOptimistic)

 If rstProjSoum.EOF Then
 bExiste = False
 Else
 bExiste = True

 Call MsgBox("Ce numéro existe dans les projets mécaniques!", vbOKOnly, "Erreur")
 End If

 Call rstProjSoum.Close
 End If
 
 'Si le projet ou la soumission n'existe pas
 If bExiste = False Then
 'Pour pouvoir réafficher l'élément affiché avant l'ajout au cas où la personne
 'annule l'ajout
 Call rstProjSoum.Open("SELECT * FROM GrbProjSoum WHERE IDProjSoum = '" & sNumero & "'", g_connData, adOpenDynamic, adLockOptimistic)

 If Not rstProjSoum.EOF Then
4 If rstProjSoum.Fields("Ouvert") = False Then
4 Call MsgBox("Ce numéro est fermé!", vbOKOnly, "Erreur")

4 Call rstProjSoum.Close
4 Set rstProjSoum = Nothing

4 Screen.MousePointer = vbDefault

4 Exit Sub
4 End If
4 End If

4 Call rstProjSoum.Close
4 Set rstProjSoum = Nothing

4 m_sAncienProjSoum = txtNoProjSoum.Text
 
4  Call InitialiserVariables(txtNoProjSoum.Text)
 
 'Affiche le nouveau numéro
4  txtNoProjSoum.Text = sNumero
 
4  If Right$(sNumero, 2) >= 60 And Right$(sNumero, 2) <=   Then
4  m_sLiaison = InputBox("Quelle est l'extention au projet " & Left$(txtNoProjSoum.Text, 6) & " auquel ce projet sera lié?", , sExtension)
4  End If
 
4  m_bModeAjout = True
 
 'Vide le listview
4  Call lvwSoumission.ListItems.Clear
 
 'On recalcul le prix
4  Call CalculerPrix
 
 'Débarre les champs
50 Call BarrerChamps(False)
 
m_sTempsDessin = "0"
 m_sTempsFabrication = "0"
 m_sTempsAssemblage = "0"
 m_sTempsProgInterface = "0"
 m_sTempsProgAutomate = "0"
 m_sTempsProgRobot = "0"
 m_sTempsVision = "0"
 m_sTempsTest = "0"
 m_sTempsInstallation = "0"
 m_sTempsMiseService = "0"
 m_sTempsFormation = "0"
5  m_sTempsGestion = "0"
5  m_sTempsShipping = "0"
 
5  m_sNbrePersonne = "0"
5  m_sTempsHebergement = "0"
5  m_sTempsRepas = "0"
5  m_sTempsTransport = "0"
5  m_sTempsUniteMobile = "0"
5  m_sPrixEmballage = "0"
 
60 txtNbreManuel.Text = "0"
  txtPrixManuel.Text = "0"

  txtForfait.Text = ""
  lblForfaitInitiale.Caption = ""

  txtPrixReception.Text = "0"
  txtPrixSoumission.Text = "0"
 
  txtPrixTotal.Text = "0"
  txtProfit.Text = "0"
  txtCommission.Text = "0"
  txtTotalTemps.Text = "0"
  txtTotalPieces.Text = "0"
  txtImprevus.Text = "0"
6  txtNoSoumission.Text = vbNullString
 
 'Vide la valeur par défaut si demande Sous-Section
6  m_sSousSection = vbNullString

6  txtProjet.Text = vbNullString
 
6  m_bModeAjout = True
6  m_bModeAffichage = False
6  m_bExtra = True
 
6  lvwSoumission.Height = lvwSoumission.Height * 0.49
6  lvwSoumission.Top = lvwPieces.Top + lvwPieces.Height + (lvwSoumission.Height * 0.02)
 
 'Met le form en mode ajout/modif
70 Call AfficherControles(MODE_AJOUT_MODIF)
  End If
  End If
 
  Screen.MousePointer = vbDefault

  Exit Sub

Oups:

  wOups "frmProjSoumElec", "cmdExtra_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdHistorique_Click()

 On Error GoTo Oups

 'Ouverture de l'historique des modifications
 If cmbProjSoum.ListCount > 0 Then
 Call RemplirListViewModifications

 lvwHistorique.Visible = True
 
 Call lvwHistorique.SetFocus
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "cmdHistorique_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdBavards_Click()

 On Error GoTo Oups

 'Ouverture de l'historique des suppressions de pièces
 If cmbProjSoum.ListCount > 0 Then
 Call RemplirListViewSuppression

 lvwBavard.Visible = True
 
 Call lvwBavard.SetFocus
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "cmdBavards_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdLegende_Click()

 On Error GoTo Oups

 Call OuvrirForm(frmLegende, True)

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "cmdLegende_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdOKDateRequise_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
 On Error GoTo Oups

 If m_bMonthViewHasFocus = True Then
 Call cmdOKDateRequise_Click
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "cmdOKDateRequise_MouseUp", Err, Err.number, Err.Description
End Sub

Private Sub cmdOKFRS_Click()

 On Error GoTo Oups

 If m_bPieceInutile = True Or m_bChangementFRS = True Then
 Call ChoisirFournisseurMateriel
 Else
 Call ChoisirFournisseur
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "cmdOKFRS_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdOKPieceTrouve_Click()

 On Error GoTo Oups

 m_bRecherchePiece = True
 m_bPieceInutile = False

 'On affiche lvwFournisseur selon la pièce choisie
 'Rempli les fournisseurs de la pièce choisie
 Call AfficherListeFournisseurs
 
 'si le listview n'est pas vide
 If lvwfournisseur.ListItems.count = 1 Then
 If MsgBox("Il n'y a aucun fournisseur pour cette pièce!" & vbNewLine & "Voulez-vous en ajouter?", vbYesNo, "Erreur") = vbYes Then
 Screen.MousePointer = vbHourglass
 
 'On ouvre le catalogue sur cet enregistrement
 Call FrmCatalogueElec.AfficherForm(lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_CATEGORIE), lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_MANUFACT), lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM))
 
 Screen.MousePointer = vbDefault
 
 'On rappelle la méthode
 Call AfficherListeFournisseurs
 End If
  End If

  Exit Sub

Oups:

  wOups "frmProjSoumElec", "cmdOKPieceTrouve_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdRafraichir_Click()

 On Error GoTo Oups

 If m_sTri <> vbNullString Then
 m_sTri = vbNullString
 
 Call RemplirListViewPieces
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "cmdRafraichir_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdRapportFACT_Click()

 On Error GoTo Oups

 Dim rstProjSoum As ADODB.Recordset
 Dim sUser As String
 Dim sNoFacture As String

 If lvwSoumission.ListItems.count > 0 Then
 If txtNoProjSoum.Text <> vbNullString Then
 If VerifierSiOuvert(sUser) = False Then
 sNoFacture = lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_FACTURATION)

 If Left$(sNoFacture, 2) = "F-" Or sNoFacture = "NC" Then
 Set rstProjSoum = New ADODB.Recordset

 'Ouvre les tables
 Call rstProjSoum.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = '" & txtNoProjSoum.Text & "' ORDER BY IDProjet", g_connData, adOpenDynamic, adLockOptimistic)

 '***********************************************************************************
 'AJOUT PAR GAÉTAN GINGRAS 0  FÉVRIER 2010
 '***********************************************************************************
 If MsgBox("Désirez-vous afficher les dates de réception et de commande?", vbYesNo, "Date de réception et de commande") = vbYes Then
 bFlag = True
 Else
 bFlag = False
 End If
 '***********************************************************************************

  Call ImprimerProjSoumFacturation(rstProjSoum, sNoFacture)
  Call ImprimerListePiecesFacturation(rstProjSoum, sNoFacture)

  'Call rstProjSoum.MoveNext

  Call rstProjSoum.Close
  Set rstProjSoum = Nothing
  Else
  Call MsgBox("La ligne sélectionnée ne contient aucune facture!", vbOKOnly, "Erreur")
  End If
 Else
 Call MsgBox("Ce projet est en modification par " & sUser & "!", vbOKOnly, "Erreur")
 End If
 End If
Else
 Call MsgBox("Ce projet ne contient aucune pièce à imprimer!", vbOKOnly, "Erreur")
End If

Exit Sub

Oups:

wOups "frmProjSoumElec", "cmdImprimer_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdRechercherClient_Click()

 On Error GoTo Oups

 Dim sRecherche As String
 
 sRecherche = InputBox("Entrez le texte à rechercher.")
 
 If StrPtr(sRecherche) <> 0 Then
 Call RemplirComboClients(sRecherche)
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "cmdRechercherClient_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdReset_Click()
 'Permet d'effacer le champs Modification et Par si c'est le user actuel
 On Error GoTo Oups

 Dim rstProjSoum As ADODB.Recordset

 If MsgBox("Êtes-vous certains de ne pas être en modification sur un autre ordinateur?", vbYesNo) = vbYes Then
 Set rstProjSoum = New ADODB.Recordset

 If m_eType = TYPE_PROJET Then
 Call rstProjSoum.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstProjSoum.Open("SELECT * FROM GrbSoumissionElec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
 End If

 rstProjSoum.Fields("Modification") = False
 rstProjSoum.Fields("Par") = ""

  Call rstProjSoum.Update

  Call rstProjSoum.Close
  Set rstProjSoum = Nothing

  cmdReset.Visible = False
  End If

  Exit Sub

Oups:

  wOups "frmProjSoumElec", "cmdReset_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdRetour_Click()

 On Error GoTo Oups

 If Right$(txtNoProjSoum.Text, 2) = "99" Then
 Call MsgBox("Vous ne pouvez pas faire de retour dans ce projet!", vbOKOnly, "Erreur")

 Exit Sub
 End If

 Screen.MousePointer = vbHourglass

 Call frmRetourMarchandise.Afficher(txtNoProjSoum.Text, g_sUserID)

 Call cmbProjSoum_Click

 Screen.MousePointer = vbDefault

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "cmdRetour_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdSortieMagasin_Click()

 On Error GoTo Oups

 Call SortieMagasin

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "cmdSortieMagasin_Click", Err, Err.number, Err.Description
End Sub

Private Sub ChangerQuantite()

 On Error GoTo Oups

 Dim sQuantite As String
 Dim itmSoum As ListItem

 sQuantite = InputBox("Quelle est la nouvelle quantité?")

 If IsNumeric(sQuantite) Then
 Set itmSoum = lvwSoumission.SelectedItem

 itmSoum.Text = sQuantite

 If Trim$(itmSoum.SubItems(I_COL_SOUM_TEMPS)) <> vbNullString Then
 'On calcul le temps * quantité pour la colonne montage
 itmSoum.SubItems(I_COL_SOUM_MONTAGE) = CDbl(Replace(itmSoum.SubItems(I_COL_SOUM_TEMPS), ".", ",")) * CDbl(Replace(itmSoum.Text, "*", vbNullString))
 Else
 itmSoum.SubItems(I_COL_SOUM_MONTAGE) = vbNullString
  End If
 
 'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
  itmSoum.SubItems(I_COL_SOUM_TOTAL) = Conversion(Round(Replace(itmSoum.Text, "*", vbNullString) * itmSoum.SubItems(I_COL_SOUM_PRIX_NET) * CSng(m_sProfit), 2), MODE_ARGENT, 2)
 
 'Pour le profit, c'est le prix total - (prix net * quantité)
  itmSoum.SubItems(I_COL_SOUM_PROFIT) = Conversion(CStr(Round(itmSoum.SubItems(I_COL_SOUM_TOTAL) - (itmSoum.SubItems(I_COL_SOUM_PRIX_NET) * Replace(itmSoum.Text, "*", vbNullString)), 2)), MODE_ARGENT)

  Call CalculerTempsFabrication

  Call CalculerPrix
  End If

  Exit Sub

Oups:

  wOups "frmProjSoumElec", "ChangerQuantite", Err, Err.number, Err.Description
End Sub

Private Sub SortieMagasin()

 On Error GoTo Oups

 Dim lColor As Long
 Dim sTag As String

 If lvwSoumission.ListItems.count > 0 Then
 If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PROFIT).Tag <> "EXTRA" Then
 'Si pas une section
 If lvwSoumission.SelectedItem.Tag <> "" Then
 'Si pas une sous-section
 If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) <> "" Then
 'Si ce n'est pas du Texte
 If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) <> "Texte" And lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) <> "Text" Then
 'Si la pièce est noire ou gris
 If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_NOIR Or lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_GRIS Then
 'Si la pièce est noire
 If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_NOIR Then
 'On la met grise
 lColor = COLOR_ORANGE

  sTag = Replace(lvwSoumission.SelectedItem.Text, "*", "")
  Else
  lColor = COLOR_NOIR

  sTag = ""
  End If
 
  lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_DESCR).ForeColor = lColor
  lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = lColor
  lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = lColor
 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = lColor
 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_MONTAGE).ForeColor = lColor
 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor = lColor
 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = lColor
 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = lColor
 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = lColor
 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_TEMPS).ForeColor = lColor
 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = lColor

 If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_FACTURATION) = "" Then
 lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_FACTURATION) = " "
 End If

 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_FACTURATION).Tag = sTag

 Call lvwSoumission.Refresh

 Call CalculerPrixReception
 End If
 End If
 End If
 End If
 Else
1  Call MsgBox("Cette commande doit être faite dans le projet " & lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PROVENANCE), vbOKOnly, "Erreur")
 End If
 End If

Exit Sub

Oups:

wOups "frmProjSoumElec", "SortieMagasin", Err, Err.number, Err.Description
End Sub

Private Sub CalculerPrixReception()

 On Error GoTo Oups

 Dim dblPrixReception As Double
 Dim iCompteur As Integer
 Dim itmProjet As ListItem

 If m_bDroitPrix = True Then
 'Pour chaque ListItems du ListView
 For iCompteur = 1 To lvwSoumission.ListItems.count
 Set itmProjet = lvwSoumission.ListItems(iCompteur)
 
 'Si ce n'est pas une section
 If itmProjet.Tag <> "" Then
 'Si ce n'est pas une sous-section
 If itmProjet.SubItems(I_COL_SOUM_PIECE) <> "" Then
 'Si c'est pas du texte
 If itmProjet.SubItems(I_COL_SOUM_PIECE) <> "Texte" And itmProjet.SubItems(I_COL_SOUM_PIECE) <> "Text" Then
 'Si c'est une réception
 If itmProjet.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_GRIS Then
  If itmProjet.ListSubItems(I_COL_SOUM_FACTURATION).Tag <> "" Then
 'On ajoute le montant
  dblPrixReception = Round(dblPrixReception + (itmProjet.SubItems(I_COL_SOUM_PRIX_NET) * Replace(itmProjet.ListSubItems(I_COL_SOUM_FACTURATION).Tag, "*", "")), 2)
  Else
  dblPrixReception = Round(dblPrixReception, 2)
  End If
  Else
 'Si c'est un retour
  If itmProjet.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_ROUGE Then
 'On soustrait le montant
  dblPrixReception = Round(dblPrixReception - (itmProjet.SubItems(I_COL_SOUM_PRIX_NET) * Replace(Replace(itmProjet.Text, "-", ""), "*", "")), 2)
 End If
 End If
 End If
 End If
 End If
 Next

 txtPrixReception.Text = Conversion(dblPrixReception, MODE_ARGENT)
Else
 txtPrixReception.Text = ""
End If

Exit Sub

Oups:

wOups "frmProjSoumElec", "CalculerPrixReception", Err, Err.number, Err.Description
End Sub

Private Sub cmdSupprimerFRS_Click()
 'Permet d'effacer un Fournisseur
 On Error GoTo Oups

 Dim sPiece As String

 'Si c'est pas "Choisir ultérieurement"
 If lvwfournisseur.SelectedItem.Text <> "CHOISIR ULTÉRIEUREMENT" Then
 If m_bPieceInutile = True Then
 sPiece = lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE)
 Else
 If m_bRecherchePiece = True Then
 sPiece = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM)
 Else
 sPiece = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_NO_ITEM)
 End If
  End If

  If MsgBox("Voulez-vous vraiment supprimer le fournisseur " & lvwfournisseur.SelectedItem.Text & " pour la pièce " & sPiece & "?", vbYesNo, "Suppression") = vbYes Then
  Call g_connData.Execute("DELETE * FROM GrbPiecesFRS WHERE NoEnreg = " & lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_ENTRER_PAR).Tag)

  Call RemplirListViewFournisseur

  frafournisseur.Visible = True

  Call lvwfournisseur.SetFocus
  End If
  End If

10 Exit Sub

Oups:

wOups "frmProjSoumElec", "cmdSupprimerFRS_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdTemps_Click()

 On Error GoTo Oups

 If cmbProjSoum.ListCount > 0 Then
 If m_eMode = MODE_AJOUT_MODIF Then
 If m_bModeAjout = True Then
 If m_bExtra = True Then
 Call frmProjSoumElecTemps.Afficher(txtNoProjSoum.Text, m_eType, m_eMode, False)
 Else
 Call frmProjSoumElecTemps.Afficher(txtNoProjSoum.Text, m_eType, m_eMode, True)
 End If
 Else
 Call frmProjSoumElecTemps.Afficher(txtNoProjSoum.Text, m_eType, m_eMode, False)
  End If
  Else
  Call frmProjSoumElecTemps.Afficher(txtNoProjSoum.Text, m_eType, m_eMode, False)
  End If
  End If

  If m_eMode = MODE_AJOUT_MODIF Then
  Call CalculerPrix
  End If
 
10 m_bTempsDejaOuvert = True

Exit Sub

Oups:

wOups "frmProjSoumElec", "cmdTemps_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdTexte_Click()

 On Error GoTo Oups

 Dim iIndex As Integer
 Dim sSousSection As String
 Dim sTexte As String

 'Ajout de texte dans la soumission
 If lvwSoumission.ListItems.count > 0 Then
 If lvwSoumission.SelectedItem.Index = 1 Then
 sSousSection = InputBox("Quelle est la sous-section?")

 If Trim$(sSousSection) = "" Then
 sSousSection = S_PAS_SOUS_SECTION
 End If

 sTexte = InputBox("Quel est le texte?")

  If Trim$(sTexte) <> "" Then
  If Len(sTexte) > 255 Then
  Call MsgBox("Le texte ne doit pas dépasser 255 caractères!", vbOKOnly, "Erreur")
  Else
  iIndex = TrouverIndexSection(sSousSection)

  Call AjouterTexte(iIndex, sTexte, sSousSection)
  End If
  End If
Else
sTexte = InputBox("Quel est le texte?")

 If Trim$(sTexte) <> "" Then
 If Len(sTexte) > 255 Then
 Call MsgBox("Le texte ne doit pas dépasser 255 caractères!", vbOKOnly, "Erreur")
 Else
 iIndex = lvwSoumission.SelectedItem.Index

 Call AjouterTexte(iIndex, sTexte, "")
 End If
 End If
 End If
End If

1  Exit Sub

Oups:

wOups "frmProjSoumElec", "cmdTexte_Click", Err, Err.number, Err.Description
End Sub

Private Sub AjouterTexte(ByVal iIndex As Integer, ByVal sTexte As String, ByVal sNomSousSection As String)

 On Error GoTo Oups

 'Méthode pour ajouter le texte
 Dim sSousSection As String
 Dim sOrdre As String
 Dim sIDSection As String

 'S'il faut l'ajouter à la fin, on prend les infos du dernier enregistrement
 If iIndex > lvwSoumission.ListItems.count Then
 If sNomSousSection = "" Then
 sSousSection = lvwSoumission.ListItems(iIndex - 1).ListSubItems(I_COL_SOUM_PIECE).Tag
 Else
 sSousSection = sNomSousSection
 End If

 sOrdre = lvwSoumission.ListItems(iIndex - 1).ListSubItems(I_COL_SOUM_MANUFACT).Tag
  sIDSection = lvwSoumission.ListItems(iIndex - 1).Tag
  Else
 'Si c'est une section
  If lvwSoumission.ListItems(iIndex).Tag = vbNullString Then
  If sNomSousSection = "" Then
  sSousSection = lvwSoumission.ListItems(iIndex - 1).ListSubItems(I_COL_SOUM_PIECE).Tag
  Else
  sSousSection = sNomSousSection
  End If

 sOrdre = lvwSoumission.ListItems(iIndex - 1).ListSubItems(I_COL_SOUM_MANUFACT).Tag
sIDSection = lvwSoumission.ListItems(iIndex - 1).Tag
 Else
 If lvwSoumission.ListItems(iIndex).SubItems(I_COL_SOUM_PIECE) = vbNullString Then
 'Si c'est pas la première sous-section
 If lvwSoumission.ListItems(iIndex - 1).Tag <> vbNullString Then
 If sNomSousSection = "" Then
 sSousSection = lvwSoumission.ListItems(iIndex - 1).ListSubItems(I_COL_SOUM_PIECE).Tag
 Else
 sSousSection = sNomSousSection
 End If

 sOrdre = lvwSoumission.ListItems(iIndex - 1).ListSubItems(I_COL_SOUM_MANUFACT).Tag
 sIDSection = lvwSoumission.ListItems(iIndex).Tag
 Else
 Call MsgBox("Vous ne pouvez pas mettre du texte entre une section et une sous-section!", vbOKOnly, "Erreur")
 
 Exit Sub
 End If
 Else
 If sNomSousSection = "" Then
 sSousSection = lvwSoumission.ListItems(iIndex).ListSubItems(I_COL_SOUM_PIECE).Tag
1  Else
 sSousSection = sNomSousSection
 End If

 sOrdre = lvwSoumission.ListItems(iIndex).ListSubItems(I_COL_SOUM_MANUFACT).Tag
 sIDSection = lvwSoumission.ListItems(iIndex).Tag
 End If
 End If
End If
 
Call lvwSoumission.ListItems.Add(iIndex)
 
Call ValeurParDefaut(lvwSoumission.ListItems(iIndex))

If m_eLangage = ANGLAIS Then
 lvwSoumission.ListItems(iIndex).SubItems(I_COL_SOUM_PIECE) = "Text"
Else
lvwSoumission.ListItems(iIndex).SubItems(I_COL_SOUM_PIECE) = "Texte"
End If

2  lvwSoumission.ListItems(iIndex).SubItems(I_COL_SOUM_DESCR) = sTexte

 'ID de la section
lvwSoumission.ListItems(iIndex).Tag = sIDSection
 
 'On ne peut pas écrire dans le tag si c'est vide
2  lvwSoumission.ListItems(iIndex).SubItems(I_COL_SOUM_MANUFACT) = " "
 
 'Ordre de la section
lvwSoumission.ListItems(iIndex).ListSubItems(I_COL_SOUM_MANUFACT).Tag = sOrdre
 
 'Nom de la sous-section
2  lvwSoumission.ListItems(iIndex).ListSubItems(I_COL_SOUM_PIECE).Tag = sSousSection

Exit Sub

Oups:

30 wOups "frmProjSoumElec", "AjouterTexte", Err, Err.number, Err.Description
End Sub

Private Sub cmdTri_Click()

 On Error GoTo Oups

 m_sTri = InputBox("Quel est le texte à trier?")
 
 m_iCol = cmbTri.ListIndex
 
 If m_sTri <> vbNullString Then
 Call RemplirListViewPieces
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "cmdTri_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdPhotos_Click()

 On Error GoTo Oups
 
 If txtCheminPhotos.Text <> vbNullString Then
 Call frmPhotoProjSoum.Afficher(txtCheminPhotos.Text)
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "cmdPhotos_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdReception_Click()

 On Error GoTo Oups

 Dim iCompteur As Integer
 Dim bOuvert As Boolean

 If Right$(txtNoProjSoum.Text, 2) = "99" Then
 Call MsgBox("Vous ne pouvez pas faire de réception pour ce projet!", vbOKOnly, "Erreur")

 Exit Sub
 End If

 For iCompteur = 0 To Forms.count - 1
 If Forms(iCompteur).Name = "FrmReceptionElec" Then
 bOuvert = True

 Exit For
  End If
  Next

  If bOuvert = True Then
  Call Unload(FrmReceptionElec)
  End If

  Call FrmReceptionElec.AfficherProjet(g_sUserID, txtNoProjSoum.Text)

  Call RemplirListViewProjSoum(txtNoProjSoum.Text)

  Exit Sub

Oups:

10 wOups "frmProjSoumElec", "cmdReception_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdBrowse_Click()

 On Error GoTo Oups

 Call OuvrirForm(frmChoixDossier, True)

 If m_bAnnulerChemin = False Then
 txtCheminPhotos.Text = m_sChemin
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "cmdBrowse_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

 On Error GoTo Oups

 If m_eMode = MODE_AJOUT_MODIF Then
 Call MsgBox("Veuillez enregistrer ou annuler avant de fermer!", vbOKOnly, "Erreur")

 Cancel = 1
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "Form_QueryUnload", Err, Err.number, Err.Description
End Sub

Private Sub Form_Resize()

 On Error GoTo Oups

 If Me.Height > I_HEIGHT_AFFICHAGE Then
 If m_eMode = MODE_INACTIF Then
 lvwPieces.width = Me.width - 375
 lvwSoumission.width = Me.width - 375
 
 lvwSoumission.Height = Me.Height - I_HEIGHT_AFFICHAGE
 lvwPieces.Height = lvwSoumission.Height * 0.49
 End If
 
 cmdImprimer.Top = Me.Height - 825
 Cmdajouter.Top = Me.Height - 825
 cmdModifier.Top = Me.Height - 825
  cmdsupprimer.Top = Me.Height - 825
  Cmdfermer.Top = Me.Height - 825
  cmdEnregistrer.Top = Me.Height - 825
  cmdAnnuler.Top = Me.Height - 825
  cmdTexte.Top = Me.Height - 825
  cmdCreerProjet.Top = Me.Height - 825
  cmdCopier.Top = Me.Height - 825
  cmdRetour.Top = Me.Height - 825
cmdBonCommande.Top = Me.Height - 825
1 cmdDemande.Top = Me.Height - 825
 cmdAnglaisFrancais.Top = Me.Height - 825
 cmdExtra.Top = Me.Height - 825
 cmdCatalogue.Top = Me.Height - 825
 cmdMaterielInutile.Top = Me.Height - 825
 cmdReset.Top = Me.Height - 825
 cmdMauvaisPrix.Top = Me.Height - 825
 cmdRapportFACT.Top = Me.Height - 825
 cmdSortieMagasin.Top = Me.Height - 825
 cmdReception.Top = Me.Height - 825
End If

1  Call PositionnerBoutons

Exit Sub

Oups:

 wOups "frmProjSoumElec", "Form_Resize", Err, Err.number, Err.Description
End Sub

Private Sub PositionnerBoutons()

 On Error GoTo Oups

 Cmdfermer.Left = Me.width - 1230
 cmdModifier.Left = Me.width - 2310
 cmdAnnuler.Left = Me.width - 2310
 cmdEnregistrer.Left = Me.width - 3390
 cmdCatalogue.Left = Me.width - 6630
 
 If m_eType = TYPE_PROJET Then
 cmdMaterielInutile.Left = Me.width - 7710
 cmdMauvaisPrix.Left = Me.width - 8790
 cmdSortieMagasin.Left = Me.width - 9870

 If m_bSupprimer = True Then
  cmdsupprimer.Left = Me.width - 3390
  Cmdajouter.Left = Me.width - 4470
  cmdBonCommande.Left = Me.width - 5550
  cmdDemande.Left = Me.width - 6630
  cmdExtra.Left = Me.width - 7710
  cmdRetour.Left = Me.width - 8790
  cmdReception.Left = Me.width - 9870
  Else
 Cmdajouter.Left = Me.width - 3390
cmdBonCommande.Left = Me.width - 4470
 cmdDemande.Left = Me.width - 5550
 cmdExtra.Left = Me.width - 6630
 cmdRetour.Left = Me.width - 7710
 cmdReception.Left = Me.width - 8790
 End If
Else
 cmdsupprimer.Left = Me.width - 3390
 Cmdajouter.Left = Me.width - 4470
 cmdCopier.Left = Me.width - 5550
 cmdMauvaisPrix.Left = Me.width - 7710
cmdDemande.Left = Me.width - 6630
 cmdCreerProjet.Left = Me.width - 7710
 End If

Exit Sub

Oups:

 wOups "frmProjSoumElec", "PositionnerBoutons", Err, Err.number, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)

 On Error GoTo Oups

 Set FrmProjSoumElec = Nothing

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "Form_Unload", Err, Err.number, Err.Description
End Sub

Private Sub lvwHistorique_LostFocus()

 On Error GoTo Oups

 'Lorsque l'historique perd le focus, on l'enlève
 lvwHistorique.Visible = False

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "lvwHistorique_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub lvwBavard_LostFocus()

 On Error GoTo Oups

 'Lorsque le bavard perd le focus, on l'enlève
 lvwBavard.Visible = False

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "lvwBavard_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub lvwfournisseur_KeyDown(KeyCode As Integer, Shift As Integer)
 'Permet d'effacer un Fournisseur
 On Error GoTo Oups

 Dim sPiece As String

 If KeyCode = vbKeyDelete Then
 If g_bModificationCatalogueElec = True Then
 'Si c'est pas Choisir Ultérieurement
 If lvwfournisseur.SelectedItem.Text <> "CHOISIR ULTÉRIEUREMENT" Then
 If m_bPieceInutile = True Then
 sPiece = lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE)
 Else
 If m_bRecherchePiece = True Then
 sPiece = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM)
 Else
  sPiece = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_NO_ITEM)
  End If
  End If

  If MsgBox("Voulez-vous vraiment supprimer le fournisseur " & lvwfournisseur.SelectedItem.Text & " pour la pièce " & sPiece & "?", vbYesNo, "Suppression") = vbYes Then
  Call g_connData.Execute("DELETE * FROM GrbPiecesFRS WHERE NoEnreg = " & lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_ENTRER_PAR).Tag)

  Call RemplirListViewFournisseur

  frafournisseur.Visible = True

  Call lvwfournisseur.SetFocus
 End If
End If
 End If
End If

Exit Sub

Oups:

wOups "frmProjSoumElec", "lvwFournisseur_KeyDown", Err, Err.number, Err.Description
End Sub

Private Sub lvwPieces_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

 On Error GoTo Oups

 Dim sTexte As String
 
 sTexte = InputBox("Quel est le texte à rechercher?")
 
 If Trim$(sTexte) <> vbNullString Then
 If Len(Trim$(sTexte)) >= 2 Then
 Call RemplirListViewRecherche(ColumnHeader.Index - 1, sTexte)

 If lvwPieceTrouve.ListItems.count > 0 Then
 fraPieceTrouve.Visible = True
 Else
 Call MsgBox("Aucun enregistrements trouvés!", vbOKOnly, "Erreur")
 End If
  Else
  Call MsgBox("Il faut un minimum de 2 caractères pour rechercher!", vbOKOnly, "Erreur")
  End If
  End If

  Exit Sub

Oups:

  wOups "frmProjSoumElec", "lvwPieces_ColumnClick", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListViewRecherche(ByVal iIndexColumn As Integer, ByVal sTexte As String)

 On Error GoTo Oups

 Dim rstPiece As ADODB.Recordset
 Dim itmPiece As ListItem
 Dim iCompteur As Integer
 Dim sChamps As String
 Dim sRecherche As String
 Dim sLettre As String

 Call lvwPieceTrouve.ListItems.Clear

 If iIndexColumn = I_COL_PIECES_NO_ITEM Then
 For iCompteur = 1 To Len(sTexte)
 sLettre = Mid$(sTexte, iCompteur, 1)

  If (Asc(sLettre) >= 4 And Asc(sLettre) <= 57) Or _
 (Asc(sLettre) >= 65 And Asc(sLettre) <= 90) Or _
 (Asc(sLettre) >=   And Asc(sLettre) <= 122) Then
  sRecherche = sRecherche & sLettre
  End If
  Next
  End If

 'Attribue le nom du champs selon la colonne cliquée
  Select Case iIndexColumn
 Case I_COL_PIECES_PIECE_GRB: sChamps = "PIECE_GRB"
  Case I_COL_PIECES_NO_ITEM: sChamps = "PIECE_MODIF"
  Case I_COL_PIECES_DESCR_EN: sChamps = "DESC_EN"
Case I_COL_PIECES_DESCR_FR: sChamps = "DESC_FR"
1 Case I_COL_PIECES_MANUFACT: sChamps = "FABRICANT"
End Select

Set rstPiece = New ADODB.Recordset
 
If iIndexColumn = I_COL_PIECES_NO_ITEM Then
 Call rstPiece.Open("SELECT * FROM GrbCatalogueElec WHERE INSTR(1," & sChamps & ",'" & sRecherche & "') > 0 ", g_connData, adOpenDynamic, adLockOptimistic)
Else
 Call rstPiece.Open("SELECT * FROM GrbCatalogueElec WHERE INSTR(1," & sChamps & ",'" & Replace(sTexte, "'", "''") & "')> 0 ", g_connData, adOpenDynamic, adLockOptimistic)
End If

 'Pour chaque enregistrement
Do While Not rstPiece.EOF
 'On ajoute dans le ListView
 Set itmPiece = lvwPieceTrouve.ListItems.Add

 If Not IsNull(rstPiece.Fields("TEMPS")) Then
 itmPiece.Tag = rstPiece.Fields("TEMPS")
 Else
 itmPiece.Tag = vbNullString
 End If

 If Not IsNull(rstPiece.Fields("PIECE_GRB")) Then
 itmPiece.Text = rstPiece.Fields("PIECE_GRB")
 Else
1  itmPiece.Text = ""
 End If

 itmPiece.SubItems(I_COL_RECH_NO_ITEM) = rstPiece.Fields("PIECE")
 itmPiece.SubItems(I_COL_RECH_CATEGORIE) = rstPiece.Fields("CATEGORIE")

 If Not IsNull(rstPiece.Fields("FABRICANT")) Then
 itmPiece.SubItems(I_COL_RECH_MANUFACT) = rstPiece.Fields("FABRICANT")
 Else
 itmPiece.SubItems(I_COL_RECH_MANUFACT) = ""
 End If

 If Not IsNull(rstPiece.Fields("DESC_EN")) Then
 itmPiece.SubItems(I_COL_RECH_DESCR_EN) = rstPiece.Fields("DESC_EN")
 Else
 itmPiece.SubItems(I_COL_RECH_DESCR_EN) = ""
End If

 If Not IsNull(rstPiece.Fields("DESC_FR")) Then
 itmPiece.SubItems(I_COL_RECH_DESCR_FR) = rstPiece.Fields("DESC_FR")
 Else
 itmPiece.SubItems(I_COL_RECH_DESCR_FR) = ""
 End If

Call rstPiece.MoveNext
Loop

30 Call rstPiece.Close

Set rstPiece = Nothing

Exit Sub

Oups:

wOups "frmProjSoumElec", "RemplirListViewRecherche", Err, Err.number, Err.Description
End Sub

Private Sub lvwPieces_KeyDown(KeyCode As Integer, Shift As Integer)

 On Error GoTo Oups

 Dim sTexte As String

 If Shift = vbCtrlMask Then
 If KeyCode = vbKeyF Then
 sTexte = InputBox("Quel est le texte à rechercher?")

 If Trim$(sTexte) <> vbNullString Then
 If Len(Trim$(sTexte)) >= 2 Then
 Call RemplirListViewRecherche(1, sTexte)

 If lvwPieceTrouve.ListItems.count > 0 Then
 fraPieceTrouve.Visible = True
 Else
  Call MsgBox("Aucun enregistrements trouvés!", vbOKOnly, "Erreur")
  End If
  Else
  Call MsgBox("Il faut un minimum de 2 caractères pour rechercher!", vbOKOnly, "Erreur")
  End If
  End If
  End If
  Else
If KeyCode = vbKeyReturn Then
Call lvwPieces_DblClick
 End If
End If

Exit Sub

Oups:

wOups "frmProjSoumElec", "lvwPieces_KeyDown", Err, Err.number, Err.Description
End Sub

Private Sub lvwPieceTrouve_DblClick()

 On Error GoTo Oups

 Dim iCompteur As Integer

 m_bRecherchePiece = True
 m_bPieceInutile = False
 m_bChangementFRS = False

 'On affiche lvwFournisseur selon la pièce choisie
 'Rempli les fournisseurs de la pièce choisie
 Call AfficherListeFournisseurs
 
 'si le listview n'est pas vide
 If lvwfournisseur.ListItems.count = 1 Then
 If MsgBox("Il n'y a aucun fournisseur pour cette pièce!" & vbNewLine & "Voulez-vous en ajouter?", vbYesNo, "Erreur") = vbYes Then
 Screen.MousePointer = vbHourglass
 
 'On ouvre le catalogue sur cet enregistrement
 Call FrmCatalogueElec.AfficherForm(lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_CATEGORIE), lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_MANUFACT), lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM))
 
 Screen.MousePointer = vbDefault
 
 'On rappelle la méthode
  Call AfficherListeFournisseurs
  End If
  End If

  Exit Sub

Oups:

  wOups "frmProjSoumElec", "lvwPieceTrouve_DblClick", Err, Err.number, Err.Description
End Sub

Private Sub lvwSoumission_DblClick()

 On Error GoTo Oups
 
 'Si il y a des enregistrements
 If lvwSoumission.ListItems.count > 0 Then
 If m_eMode = MODE_AJOUT_MODIF Then
 'Si c'est une sous-section
 If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) = vbNullString Then
 Call ModifierSousSection
 Else
 'Si c'est une pièce
 If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) <> vbNullString And lvwSoumission.SelectedItem.Tag <> vbNullString Then
 'Si la pièce n'est pas un Text ou Texte
 If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) <> "Texte" And lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) <> "Text" Then
 'Si la pièce n'a pas de fournisseur
 If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_DISTRIB) = vbNullString Then
 If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PROFIT) = "" Then
 lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PROFIT) = " "
  End If

  If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PROFIT).Tag <> "EXTRA" Then
  Call AjouterPrix
  Else
  Call MsgBox("Cette commande doit être faite dans le projet " & lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PROVENANCE), vbOKOnly, "Erreur")
  End If
  Else
 'Si la pièce est en commande
  If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_ORANGE Then
 If MsgBox("Voulez-vous annuler cette commande?", vbYesNo) = vbYes Then
 Call AnnulerCommande
 End If
 Else
 If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_BRUN Then
 Call ChangerFournisseurRetour
 End If
 End If
 End If
 Else
 Call ModifierTexte
 End If
 End If
 End If
 End If
End If

 Exit Sub

Oups:

wOups "frmProjSoumElec", "lvwSoumission_DblClick", Err, Err.number, Err.Description
End Sub

Private Sub AjouterPrix()

 On Error GoTo Oups

 Call ViderChamps_frs

 'Rempli le combo des fournisseurs
 Call RemplirComboFournisseur

 cmbfrs.Locked = False

 m_bMauvaisPrix = False

 'Positionne le frame
 fraPrixPiece.Top = lvwSoumission.Top + 500
 
 'Montre le frame
 fraPrixPiece.Visible = True

 'Met le numéro de la pièce dans le tag
 fraPrixPiece.Tag = lvwSoumission.SelectedItem.Index
 
 'Donne le focus au combo
 Call cmbfrs.SetFocus

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "AjouterPrix", Err, Err.number, Err.Description
End Sub

Private Sub ModifierTexte()

 On Error GoTo Oups

 Dim sTexte As String

 sTexte = InputBox("Quel est le nouveau texte?", , lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_DESCR))

 If sTexte <> "" Then
 If Len(sTexte) > 255 Then
 Call MsgBox("Le texte ne pas dépasser 255 caractères!", vbOKOnly, "Erreur")
 Else
 lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_DESCR) = sTexte
 End If
 End If

 Exit Sub

Oups:

  wOups "frmProjSoumElec", "ModifierTexte", Err, Err.number, Err.Description
End Sub

Private Sub ModifierSousSection()

 On Error GoTo Oups

 Dim sSousSection As String
 Dim sAncienneSS As String
 Dim sTag As String
 Dim iCompteur As Integer

 sSousSection = InputBox("Quel est le nouveau nom de la sous-section?", , lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_DESCR))

 If StrPtr(sSousSection) <> 0 Then
 sAncienneSS = lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_DESCR)
 
 If sAncienneSS = vbNullString Then
 sAncienneSS = S_PAS_SOUS_SECTION
 End If
 
  If Trim$(sSousSection) = vbNullString Then
  sTag = S_PAS_SOUS_SECTION
  sSousSection = vbNullString
  Else
  sTag = sSousSection
  End If

  lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_DESCR) = sSousSection
  lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_DESCR).Bold = True
 
For iCompteur = lvwSoumission.SelectedItem.Index + 1 To lvwSoumission.ListItems.count
If lvwSoumission.ListItems(iCompteur).ListSubItems(I_COL_SOUM_PIECE).Tag = sAncienneSS Then
 lvwSoumission.ListItems(iCompteur).ListSubItems(I_COL_SOUM_PIECE).Tag = sTag
 Else
 Exit For
 End If
 Next
End If

Exit Sub

Oups:

wOups "frmProjSoumElec", "ModifierSousSection", Err, Err.number, Err.Description
End Sub

Private Sub ChangerFournisseurRetour()

 On Error GoTo Oups

' m_bPieceInutile = True
 m_bRecherchePiece = False
 m_bChangementFRS = True

 Call AfficherListeFournisseurs

 If lvwfournisseur.ListItems.count = 0 Then
 Call MsgBox("Il n'y a aucun fournisseur pour cette pièce!", vbOKOnly, "Erreur")
 
 Exit Sub
 Else
 frafournisseur.Visible = True
 End If

  Exit Sub

Oups:

  wOups "frmProjSoumElec", "ChangerFournisseurRetour", Err, Err.number, Err.Description
End Sub

Private Sub lvwSoumission_ItemCheck(ByVal Item As MSComctlLib.ListItem)

 On Error GoTo Oups

 'Si le programme n'est pas en mode ajout ou modif
 If m_eMode = MODE_INACTIF Then
 'On annule le ItemCheck
 Item.Checked = Not Item.Checked
 Else
 'Si c'est une section, une sous-section ou du texte
 If Item.Tag = vbNullString Or Item.SubItems(I_COL_SOUM_PIECE) = vbNullString Then
 'On annule le ItemCheck
 Item.Checked = Not Item.Checked
 End If
 End If

 Exit Sub

Oups:

  wOups "frmProjSoumElec", "lvwSoumission_ItemCheck", Err, Err.number, Err.Description
End Sub

Private Sub lvwSoumission_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
 On Error GoTo Oups

 Dim iNbreSelected As Integer
 Dim iIndexSelected As Integer
 Dim iCompteur As Integer
 Dim bAfficherMenu As Boolean

 If m_eMode = MODE_AJOUT_MODIF Then
 If Button = vbRightButton Then
 If lvwSoumission.ListItems.count > 0 Then
 'S'il y a plusieurs items de sélectionnés, c'est parce que l'utilisateur
 'a sélectionné plusieurs items
 'Donc, on ne désélectionne pas
 For iCompteur = 1 To lvwSoumission.ListItems.count
 If lvwSoumission.ListItems(iCompteur).Selected = True Then
 iNbreSelected = iNbreSelected + 1

  iIndexSelected = iCompteur
  End If
  Next

  If iNbreSelected = 1 Then
  lvwSoumission.ListItems(iIndexSelected).Selected = False
  End If

  Set lvwSoumission.DropHighlight = lvwSoumission.HitTest(X, Y)

  If Not lvwSoumission.DropHighlight Is Nothing Then
 If iNbreSelected = 1 Then
 lvwSoumission.DropHighlight.Selected = True

 If lvwSoumission.SelectedItem.Tag <> "" Then
 If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PROFIT) = "" Then
 lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PROFIT) = " "
 End If

 If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PROFIT).Tag <> "EXTRA" Then
 bAfficherMenu = True
 Else
 bAfficherMenu = False
 End If
 End If
 Else
 If lvwSoumission.DropHighlight.Selected = True Then
 If lvwSoumission.SelectedItem.Tag <> "" Then
 If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PROFIT).Tag <> "EXTRA" Then
 If g_bModificationFacturation = True Then
 bAfficherMenu = True
 Else
1  bAfficherMenu = False
 End If
 Else
 bAfficherMenu = False
 End If
 Else
 bAfficherMenu = False
 End If
 End If
 End If
 Else
 bAfficherMenu = False
 End If

 If bAfficherMenu = True Then
 Call RemplirOptionsMenuRightClick(iNbreSelected)

 Call PopupMenu(mnuRightClick)
 End If
 End If
 Else
 If Shift <> vbCtrlMask And Shift <> vbShiftMask Then
 Set lvwSoumission.DropHighlight = Nothing
 End If
3 End If
End If

Exit Sub

Oups:

wOups "frmProjSoumElec", "lvwSoumission_MouseDown", Err, Err.number, Err.Description
End Sub

Private Sub RemplirOptionsMenuRightClick(ByVal iNbreSelected As Integer)

 On Error GoTo Oups

 Dim bFacturer As Boolean
 Dim bNC As Boolean
 Dim bDateRequise As Boolean
 Dim bCommentaire As Boolean
 Dim bID As Boolean
 Dim bMauvaisPrix As Boolean
 Dim bMaterielInutile As Boolean
 Dim bTexte As Boolean
 Dim bSousSection As Boolean
 Dim bFournisseur As Boolean
  Dim bAnnulerCommande As Boolean
  Dim bSupprimer As Boolean
  Dim bAjouterPrix As Boolean
  Dim bSortieMagasin As Boolean
  Dim bChangerQuantite As Boolean

  If iNbreSelected > 1 Then
  If m_eType = TYPE_PROJET Then
  If g_bModificationFacturation = True Then
 bFacturer = True
 bNC = True
 bSupprimer = True
 End If
 Else
 bSupprimer = True
 End If
Else
 'Si c'est une sous-section
 If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) = "" Then
 bSousSection = True
 Else
 If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) = "Texte" Or lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) = "Text" Then
 bTexte = True

 If (m_eType = TYPE_SOUMISSION) Or ((m_eType = TYPE_PROJET) And (Right$(txtNoProjSoum.Text, 2) > 19)) Then
 bSupprimer = True
 End If
 Else
 If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor = -2147483640 Then
 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor = 0
1  End If

 Select Case lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor
 Case COLOR_ORANGE:
 If g_bModificationFacturation = True Then
 bFacturer = True
 bNC = True
 End If

 bID = True
 bDateRequise = True
 bCommentaire = True
 bAnnulerCommande = True
 bMauvaisPrix = True
 
 Case COLOR_BRUN:
 bCommentaire = True
 bFournisseur = True

 If (m_eType = TYPE_SOUMISSION) Or ((m_eType = TYPE_PROJET) And (Right$(txtNoProjSoum.Text, 2) > 19)) Then
 bSupprimer = True
 End If

 Case COLOR_GRIS:
 If g_bModificationFacturation = True Then
 bFacturer = True
 bNC = True
 End If

 bCommentaire = True
 bID = True
 bMauvaisPrix = True
 bMaterielInutile = True

 Case COLOR_VERT_FORET:
 bCommentaire = True

 If (m_eType = TYPE_SOUMISSION) Or ((m_eType = TYPE_PROJET) And (Right$(txtNoProjSoum.Text, 2) > 19)) Then
 bSupprimer = True
 End If

 Case COLOR_ROUGE:
 bCommentaire = True

 Case COLOR_MAGENTA:
 bCommentaire = True
 bAjouterPrix = True

 If m_eType = TYPE_PROJET Then
 bID = True
 End If

 If m_eType = TYPE_SOUMISSION Then
 bChangerQuantite = True
 End If

 Case COLOR_NOIR:
 If m_eType = TYPE_PROJET Then
 If g_bModificationFacturation = True Then
 bFacturer = True
 bNC = True
 End If

4 bID = True
4 bMaterielInutile = True
4 bSortieMagasin = True
4 Else
4 bChangerQuantite = True
4 End If

4 bCommentaire = True
4 bMauvaisPrix = True
4 bFournisseur = True
4 bSupprimer = True
4 End Select
4  End If
4  End If
4  End If

 'Pour empeche que tous les éléments deviennent invisible, je les mets visible au
 'début
4  mnuFacturer.Visible = True
4  mnuNC.Visible = True
4  mnuDateRequise.Visible = True
4  mnuCommentaire.Visible = True
4  mnuID.Visible = True
50 mnuMauvaisPrix.Visible = True
50 mnuInutile.Visible = True
 mnuTexte.Visible = True
 mnuChangerSS.Visible = True
 mnuFournisseur.Visible = True
 mnuAnnulerCommande.Visible = True
 mnuSupprimer.Visible = True
 mnuAjouterPrix.Visible = True
 mnuSortieMagasin.Visible = True
 mnuQuantite.Visible = True

 mnuFacturer.Visible = bFacturer
 mnuNC.Visible = bNC
5  mnuDateRequise.Visible = bDateRequise
5  mnuCommentaire.Visible = bCommentaire
5  mnuID.Visible = bID
5  mnuMauvaisPrix.Visible = bMauvaisPrix
5  mnuInutile.Visible = bMaterielInutile
5  mnuTexte.Visible = bTexte
5  mnuChangerSS.Visible = bSousSection
5  mnuFournisseur.Visible = bFournisseur
60 mnuAnnulerCommande.Visible = bAnnulerCommande
60 mnuSupprimer.Visible = bSupprimer
  mnuAjouterPrix.Visible = bAjouterPrix
  mnuSortieMagasin.Visible = bSortieMagasin
  mnuQuantite.Visible = bChangerQuantite

  Exit Sub

Oups:

  wOups "frmProjSoumElec", "RemplirOptionsMenuRightClick", Err, Err.number, Err.Description
End Sub

Private Sub mnuAjouterPrix_Click()

 On Error GoTo Oups

 Call AjouterPrix

 Call EnleverSelection

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "mnuAjouterPrix_Click", Err, Err.number, Err.Description
End Sub

Private Sub mnuAnnulerCommande_Click()

 On Error GoTo Oups

 Call AnnulerCommande

 Call EnleverSelection

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "mnuAnnulerCommande_Click", Err, Err.number, Err.Description
End Sub

Private Sub mnuChangerSS_Click()

 On Error GoTo Oups

 Call ModifierSousSection

 Call EnleverSelection

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "mnuChangerSS_Click", Err, Err.number, Err.Description
End Sub

Private Sub mnuDateRequise_Click()

 On Error GoTo Oups

 If Trim$(lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_DATE_REQUISE)) = "" Then
 mvwDateRequise.Year = Year(Date)
 mvwDateRequise.Month = Month(Date)
 mvwDateRequise.Day = Day(Date)
 Else
 mvwDateRequise.Year = Left$(lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_DATE_REQUISE), 4)
 mvwDateRequise.Month = Mid$(lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_DATE_REQUISE), 6, 2)
 mvwDateRequise.Day = Right$(lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_DATE_REQUISE), 2)
 End If

 fraDateRequise.Top = lvwSoumission.Top

  fraDateRequise.Visible = True

  Exit Sub

Oups:

  wOups "frmProjSoumElec", "mnuDateRequise_Click", Err, Err.number, Err.Description
End Sub

Private Sub mnuCommentaire_Click()

 On Error GoTo Oups

 txtcommentaire.Text = lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_COMMENTAIRE)

 fraCommentaire.Top = lvwSoumission.Top

 fraCommentaire.Visible = True

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "mnuCommentaire_Click", Err, Err.number, Err.Description
End Sub

Private Sub mnuFacturer_Click()

 On Error GoTo Oups

 Call FacturerDate

 Call EnleverSelection

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "mnuFacturer_Click", Err, Err.number, Err.Description
End Sub

Private Sub mnuFournisseur_Click()

 On Error GoTo Oups

 Call ChangerFournisseurRetour

 Call EnleverSelection

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "mnuFournisseur_Click", Err, Err.number, Err.Description
End Sub

Private Sub mnuID_Click()

 On Error GoTo Oups

 lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_ID) = InputBox("Quel est l'ID", , lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_ID))

 Call EnleverSelection

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "mnuID_Click", Err, Err.number, Err.Description
End Sub

Private Sub mnuInutile_Click()

 On Error GoTo Oups

 Call MaterielInutile

 Call EnleverSelection

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "mnuInutile_Click", Err, Err.number, Err.Description
End Sub

Private Sub mnuMauvaisPrix_Click()

 On Error GoTo Oups

 Call MauvaisPrix

 Call EnleverSelection

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "mnuMauvaisPrix", Err, Err.number, Err.Description
End Sub

Private Sub mnuNC_Click()

 On Error GoTo Oups

 Call FacturerNC

 Call EnleverSelection

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "mnuNC_Click", Err, Err.number, Err.Description
End Sub

Private Sub mnuQuantite_Click()

 On Error GoTo Oups

 Call ChangerQuantite

 Call EnleverSelection

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "mnuQuantite_Click", Err, Err.number, Err.Description
End Sub

Private Sub mnuSortieMagasin_Click()

 On Error GoTo Oups

 Call SortieMagasin

 Call EnleverSelection

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "mnuSortieMagasin_Click", Err, Err.number, Err.Description
End Sub

Private Sub mnuSupprimer_Click()
 
 On Error GoTo Oups

 Call EffacerItemListViewSoumission

 Call EnleverSelection

 Exit Sub
 
Oups:

 wOups "frmProjSoumElec", "mnuSupprimer_Click", Err, Err.number, Err.Description
End Sub

Private Sub mnuTexte_Click()

 On Error GoTo Oups

 Call ModifierTexte

 Call EnleverSelection

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "mnuTexte_Click", Err, Err.number, Err.Description
End Sub

Private Sub EnleverSelection()

 On Error GoTo Oups

 Dim iCompteur As Integer

 Set lvwSoumission.DropHighlight = Nothing

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "EnleverSelection", Err, Err.number, Err.Description
End Sub

Private Sub mvwDate_LostFocus()

 On Error GoTo Oups
 
 'Quand le calendrier perd le focus, il faut l'enlever
 mvwDate.Visible = False

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "mvwDate_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mvwDateFacturation_LostFocus()

 On Error GoTo Oups

 'Quand le calendrier perd le focus, il faut l'enlever
 mvwDateFacturation.Visible = False

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "mvwDateFacturation_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mvwDate_DateClick(ByVal DateClicked As Date)

 On Error GoTo Oups

 'Affiche la date dans le TextBox sous le format AAAA-MM-JJ
 txtDelais.Text = ConvertDate(DateClicked)
 
 'Enlever le calendrier
 mvwDate.Visible = False

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "mvwDate_DateClick", Err, Err.number, Err.Description
End Sub

Private Sub mvwDateFacturation_DateClick(ByVal DateClicked As Date)

 On Error GoTo Oups

 'Affiche la date dans le TextBox sous le format AAAA-MM-JJ
 txtDateFacturation.Text = ConvertDate(DateClicked)

 'Enlever le calendrier
 mvwDateFacturation.Visible = False

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "mvwDateFacturation_DateClick", Err, Err.number, Err.Description
End Sub

Private Sub cmdEnregistrer_Click()

 On Error GoTo Oups

 Dim objControl As Control
 Dim sMessage As String

 frafournisseur.Visible = False
 fraPieceTrouve.Visible = False
 fraCommentaire.Visible = False
 fraDateRequise.Visible = False
 
 'Vérification des textbox
 For Each objControl In Me
 If TypeOf objControl Is TextBox Then
 If objControl.Visible = True Then
 If objControl.Name <> "txtNoSoumission" And _
 objControl.Name <> "txtCheminPhotos" And _
 objControl.Name <> "txtPrixReception" And _
 objControl.Name <> "txtDateFacturation" And _
 objControl.Name <> "txtPrixSoumission" And _
 objControl.Name <> "txtDelais" And _
 objControl.Name <> "txtForfait" Then
  If Trim$(objControl.Text) = vbNullString Then
  Call MsgBox("Vous devez remplir tous les champs!", vbOKOnly, "Erreur")
 
  Exit Sub
  End If
  End If
  End If
  Else
  If TypeOf objControl Is ComboBox Then
 If objControl.Visible = True Then
 If objControl.ListIndex = -1 Then
 If objControl.Name <> "cmbTri" And _
 objControl.Name <> "cmbSections" And _
 objControl.Name <> "cmbPieces" Then
 Call MsgBox("Vous devez remplir tous les champs!", vbOKOnly, "Erreur")

 Exit Sub
 End If
 End If
 End If
 End If
 End If
Next
 
 'Vérification du transport
If cmbtransport.ListIndex = -1 Then
Call MsgBox("Vous devez choisir le transport!", vbOKOnly, "Erreur")
 
 Exit Sub
 End If

If m_eType = TYPE_SOUMISSION Then
 If m_sTempsTest = 0 Or m_sTempsDessin = 0 Then
 If MsgBox("Les temps de dessin ou de test sont vides" & vbNewLine & "Voulez - vous l'enregistrer?", vbYesNo) = vbNo Then
 Exit Sub
1  End If
 End If
 End If

Screen.MousePointer = vbHourglass

If BackupPieces(txtNoProjSoum.Text) = False Then
 If m_eType = TYPE_PROJET Then
 sMessage = "Une erreur est survenue lors de la copie de sauvegarde du projet en cours!"
 Else
 sMessage = "Une erreur est survenue lors de la copie de sauvegarde de la soumission en cours!"
 End If

 sMessage = sMessage & vbNewLine & vbNewLine & "Voulez-vous continuer ?"

 Screen.MousePointer = vbDefault

 If MsgBox(sMessage, vbYesNo) = vbNo Then
 Exit Sub
 Else
 Screen.MousePointer = vbHourglass
 End If
2  End If
 
 'Enregistre la soumission
Call EnregistrerProjSoum(txtNoProjSoum.Text)
 
2  Call OuvrirProjSoum(False)
 
 'Remet en mode inactif
Call AfficherControles(MODE_INACTIF)

30 m_bEnregistrement = True
 
 'Affiche la soumission actuel
Call AfficherProjSoum(txtNoProjSoum.Text)

m_bEnregistrement = False
 
Screen.MousePointer = vbDefault

Exit Sub

Oups:

wOups "frmProjSoumElec", "cmdEnregistrer_Click", Err, Err.number, Err.Description
End Sub

Private Sub EnregistrerFACT(ByVal sNoProjet As String)
 'Calcul le total de chaque facture dans le projet
 On Error GoTo Oups

 Dim rstModif As ADODB.Recordset
 Dim rstEmploye As ADODB.Recordset
 Dim sPrixTotal As String
 Dim sProfit As String
 Dim sCommission As String
 Dim sNoFacture As String
 Dim sTempsFab As String
 Dim sTotalPiece As String
 Dim sImprevue As String
 Dim sTotalTemps As String
  Dim sManuel As String
  Dim iCompteur As Integer
  Dim iIndexFacture As Integer
  Dim collFacture As Collection
  Dim bExiste As Boolean

  Set collFacture = New Collection

  Call g_connData.Execute("DELETE * FROM GrbProjet_Modif WHERE IDProjet = '" & sNoProjet & "' AND Type = 'E' AND TypeModif = 'FACTURATION'")

  If lvwSoumission.ListItems.count > 0 Then
Set rstModif = New ADODB.Recordset
1 Set rstEmploye = New ADODB.Recordset

 For iCompteur = 1 To lvwSoumission.ListItems.count
 bExiste = False

 sNoFacture = lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_FACTURATION)

 If Trim$(sNoFacture) <> "" Then
 For iIndexFacture = 1 To collFacture.count
 If collFacture(iIndexFacture) = sNoFacture Then
 bExiste = True

 Exit For
 End If
 Next

 If bExiste = False Then
 Call collFacture.Add(sNoFacture)

 Call CalculerPrixFacturation(sNoFacture, sCommission, sPrixTotal, sProfit, sTempsFab, sTotalPiece, sImprevue, sTotalTemps, sManuel)

 Call rstModif.Open("SELECT * FROM GrbProjet_Modif WHERE [Date] = '" & Replace(sNoFacture, "F-", "") & "' AND TypeModif = 'FACTURATION'", g_connData, adOpenDynamic, adLockOptimistic)

 If rstModif.EOF Then
 Call rstModif.AddNew
 End If

1  rstModif.Fields("IDProjet") = txtNoProjSoum.Text

 Call rstEmploye.Open("SELECT * FROM GrbEmployés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)

 rstModif.Fields("NoEmployé") = rstEmploye.Fields("NoEmploye")

 Call rstEmploye.Close

 rstModif.Fields("Date") = Replace(sNoFacture, "F-", "")
 rstModif.Fields("Heure") = " "
 rstModif.Fields("Type") = "E"
 rstModif.Fields("TypeModif") = "FACTURATION"
 rstModif.Fields("Valeur") = sPrixTotal

 Call rstModif.Update

 Call rstModif.Close
 End If
 End If
Next

 Set rstModif = Nothing
Set rstEmploye = Nothing
End If

2  Exit Sub

Oups:

wOups "frmProjSoumElec", "EnregistrerFACT", Err, Err.number, Err.Description
End Sub

Private Function BackupPieces(ByVal sNoProjSoum As String) As Boolean

 On Error GoTo Oups

 Dim rstProjSoum As ADODB.Recordset
 Dim rstProjSoumBackup As ADODB.Recordset
 Dim sDateCopie As String

 Set rstProjSoum = New ADODB.Recordset
 Set rstProjSoumBackup = New ADODB.Recordset

 If m_eType = TYPE_PROJET Then
 Call rstProjSoum.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & sNoProjSoum & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

 Call rstProjSoumBackup.Open("SELECT * FROM GrbProjet_Pieces_Tampon", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstProjSoum.Open("SELECT * FROM GrbSoumission_Pieces WHERE IDSoumission = '" & sNoProjSoum & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

  Call rstProjSoumBackup.Open("SELECT * FROM GrbSoumission_Pieces_Tampon", g_connData, adOpenDynamic, adLockOptimistic)
  End If

  sDateCopie = ConvertDate(Date) & " " & Time

  Do While Not rstProjSoum.EOF
  Call rstProjSoumBackup.AddNew

  rstProjSoumBackup.Fields("DateCopie") = sDateCopie

  If m_eType = TYPE_PROJET Then
  rstProjSoumBackup.Fields("IDProjet") = rstProjSoum.Fields("IDProjet")
Else
rstProjSoumBackup.Fields("IDSoumission") = rstProjSoum.Fields("IDSoumission")
 End If

 rstProjSoumBackup.Fields("Initiales") = g_sInitiale
 rstProjSoumBackup.Fields("IDSection") = rstProjSoum.Fields("IDSection")
 rstProjSoumBackup.Fields("NumItem") = rstProjSoum.Fields("NumItem")
 rstProjSoumBackup.Fields("Qté") = rstProjSoum.Fields("Qté")
 rstProjSoumBackup.Fields("Desc_FR") = rstProjSoum.Fields("Desc_FR")
 rstProjSoumBackup.Fields("Desc_EN") = rstProjSoum.Fields("Desc_EN")
 rstProjSoumBackup.Fields("Manufact") = rstProjSoum.Fields("Manufact")
 rstProjSoumBackup.Fields("Prix_list") = rstProjSoum.Fields("Prix_list")
 rstProjSoumBackup.Fields("Escompte") = rstProjSoum.Fields("Escompte")
rstProjSoumBackup.Fields("Prix_net") = rstProjSoum.Fields("Prix_net")
 rstProjSoumBackup.Fields("IDFRS") = rstProjSoum.Fields("IDFRS")
 rstProjSoumBackup.Fields("Temps") = rstProjSoum.Fields("Temps")
 rstProjSoumBackup.Fields("Temps_total") = rstProjSoum.Fields("Temps_total")
 rstProjSoumBackup.Fields("Prix_total") = rstProjSoum.Fields("Prix_total")
 rstProjSoumBackup.Fields("Profit_Argent") = rstProjSoum.Fields("Profit_Argent")
 rstProjSoumBackup.Fields("sousSection") = rstProjSoum.Fields("sousSection")
1  rstProjSoumBackup.Fields("OrdreSection") = rstProjSoum.Fields("OrdreSection")
 rstProjSoumBackup.Fields("NuméroLigne") = rstProjSoum.Fields("NuméroLigne")
 rstProjSoumBackup.Fields("PrixOrigine") = rstProjSoum.Fields("PrixOrigine")
 rstProjSoumBackup.Fields("Type") = rstProjSoum.Fields("Type")
 rstProjSoumBackup.Fields("Visible") = rstProjSoum.Fields("Visible")
 rstProjSoumBackup.Fields("Commandé") = rstProjSoum.Fields("Commandé")
 rstProjSoumBackup.Fields("Quoté") = rstProjSoum.Fields("Quoté")
 rstProjSoumBackup.Fields("Recu") = rstProjSoum.Fields("Recu")
 rstProjSoumBackup.Fields("Retour") = rstProjSoum.Fields("Retour")
 rstProjSoumBackup.Fields("CommandeAnnulée") = rstProjSoum.Fields("CommandeAnnulée")
 rstProjSoumBackup.Fields("ID") = rstProjSoum.Fields("ID")
 rstProjSoumBackup.Fields("PieceExtra") = rstProjSoum.Fields("PieceExtra")
 rstProjSoumBackup.Fields("PieceExtraChargeable") = rstProjSoum.Fields("PieceExtraChargeable")
rstProjSoumBackup.Fields("PieceExtraNonChargeable") = rstProjSoum.Fields("PieceExtraNonChargeable")
 rstProjSoumBackup.Fields("MatérielInutile") = rstProjSoum.Fields("MatérielInutile")
rstProjSoumBackup.Fields("Commentaire") = rstProjSoum.Fields("Commentaire")
 rstProjSoumBackup.Fields("Devise") = rstProjSoum.Fields("Devise")

If m_eType = TYPE_PROJET Then
 rstProjSoumBackup.Fields("NoRetour") = rstProjSoum.Fields("NoRetour")
 rstProjSoumBackup.Fields("DateRéception") = rstProjSoum.Fields("DateRéception")
 rstProjSoumBackup.Fields("QuantitéRecue") = rstProjSoum.Fields("QuantitéRecue")
 rstProjSoumBackup.Fields("Facturation") = rstProjSoum.Fields("Facturation")
rstProjSoumBackup.Fields("DateCommande") = rstProjSoum.Fields("DateCommande")
 rstProjSoumBackup.Fields("DateRequise") = rstProjSoum.Fields("DateRequise")
 rstProjSoumBackup.Fields("NomCommande") = rstProjSoum.Fields("NomCommande")
 rstProjSoumBackup.Fields("NoSéquentiel") = rstProjSoum.Fields("NoSéquentiel")
 rstProjSoumBackup.Fields("DateRetour") = rstProjSoum.Fields("DateRetour")
 End If

 rstProjSoumBackup.Fields("Provenance") = rstProjSoum.Fields("Provenance")

 Call rstProjSoumBackup.Update

 Call rstProjSoum.MoveNext
Loop

Call rstProjSoum.Close
3  Set rstProjSoum = Nothing

Call rstProjSoumBackup.Close
3  Set rstProjSoumBackup = Nothing

BackupPieces = True

3  Exit Function

Oups:

wOups "frmProjSoumElec", "BackupPieces", Err, Err.number, Err.Description
End Function

Private Sub EnregistrerProjSoum(ByVal sNoProjSoum As String)

 On Error GoTo Oups

 Dim rstProjSoum As ADODB.Recordset
 Dim rstPiece As ADODB.Recordset
 Dim rstEmploye As ADODB.Recordset
 Dim rstModif As ADODB.Recordset
 Dim rstOuvert As ADODB.Recordset
 Dim rstSection As ADODB.Recordset
 Dim rstPunch As ADODB.Recordset
 Dim iCompteur As Integer
 Dim itmPiece As ListItem
 Dim sTable As String
  Dim sTableModif As String
  Dim sTablePiece As String
  Dim sChamps As String
  Dim sSection As String
  Dim sExtra As String
  Dim bCalculExtra As Boolean
  Dim collExtra As Collection
  Dim iCompteurExtra As Integer
10 Dim bExiste As Boolean
Dim bAjoutCommande As Boolean
Dim dblNbrePers As Double
Dim dblJoursHebergement As Double
Dim dblJoursRepas As Double
Dim dblHebergement As Double
Dim dblHebergement As Double
Dim dblRepas As Double
Dim dblTotalHebergement As Double
 
Set rstProjSoum = New ADODB.Recordset
Set rstPiece = New ADODB.Recordset
Set rstEmploye = New ADODB.Recordset
1  Set rstModif = New ADODB.Recordset
Set rstOuvert = New ADODB.Recordset
 Set rstSection = New ADODB.Recordset

Set collExtra = New Collection

 Call rstEmploye.Open("SELECT noEmploye FROM GrbEmployés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)

 'Si c'est un projet
If m_eType = TYPE_PROJET Then
 sTable = "GrbProjetElec"
1  sTableModif = "GrbProjet_Modif"
 sTablePiece = "GrbProjet_Pieces"
 sChamps = "IDProjet"
Else
 sTable = "GrbSoumissionElec"
 sTableModif = "GrbSoumission_Modif"
 sTablePiece = "GrbSoumission_Pieces"
 sChamps = "IDSoumission"
End If

 'Si c'est un ajout
If m_bModeAjout = True Then
 'On ouvre le recordset selon le type
 Call rstProjSoum.Open("SELECT * FROM " & sTable & " WHERE " & sChamps & " = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)

 If m_eType = TYPE_PROJET Then
 If rstProjSoum.EOF Then
 bAjoutCommande = True
 Else
 bAjoutCommande = False
 End If
Else
 bAjoutCommande = False
End If

 Call rstProjSoum.AddNew

If m_eType = TYPE_PROJET Then
rstProjSoum.Fields("LiaisonChargeable") = m_sLiaison
 End If

 rstProjSoum.Fields("Creer") = ConvertDate(Date)
 rstProjSoum.Fields("Creer_Par") = rstEmploye.Fields("noEmploye")

 rstProjSoum.Fields(sChamps) = sNoProjSoum

 If m_eType = TYPE_PROJET Then
 rstProjSoum.Fields("IDSoumission") = txtNoSoumission.Text
 End If

 Call rstOuvert.Open("SELECT * FROM GrbProjSoum WHERE IDProjSoum = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)

 If rstOuvert.EOF Then
 Call rstOuvert.AddNew

 rstOuvert.Fields("IDProjSoum") = sNoProjSoum
 rstOuvert.Fields("NoClient") = cmbclient.ItemData(cmbclient.ListIndex)
 rstOuvert.Fields("Description") = txtProjet.Text
 rstOuvert.Fields("DateOuverture") = ConvertDate(Date)
 rstOuvert.Fields("Ouvert") = True
 
 If m_eType = TYPE_PROJET Then
 rstOuvert.Fields("Type") = "P"
 Else
 rstOuvert.Fields("Type") = "S"
4 End If

4 Call rstOuvert.Update
4 End If
 
4 Call rstOuvert.Close
4 Set rstOuvert = Nothing

4 m_bModeAjout = False
4 Else
4 Call EnregistrerSuppression

4 Call rstProjSoum.Open("SELECT * FROM " & sTable & " WHERE " & sChamps & " = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
4 Call rstModif.Open("SELECT * FROM " & sTableModif, g_connData, adOpenDynamic, adLockOptimistic)
 
4 Call rstModif.AddNew
 
4  rstModif.Fields("Type") = "E"
4  rstModif.Fields(sChamps) = sNoProjSoum
4  rstModif.Fields("noEmployé") = rstEmploye.Fields("noEmploye")
4  rstModif.Fields("Date") = ConvertDate(Date)
4  rstModif.Fields("Heure") = Time
4  rstModif.Fields("TypeModif") = "MODIFICATION"
 
4  Call rstModif.Update
 
4  Call rstModif.Close
50 Set rstModif = Nothing

5 Call rstOuvert.Open("SELECT * FROM GrbProjSoum WHERE IDProjSoum = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 If rstOuvert.Fields("NoClient") <> cmbclient.ItemData(cmbclient.ListIndex) Then
 rstOuvert.Fields("NoClient") = cmbclient.ItemData(cmbclient.ListIndex)

 Set rstPunch = New ADODB.Recordset

 Call rstPunch.Open("SELECT * FROM GrbPunch WHERE NoProjet = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)

 If Not rstPunch.EOF Then
 If MsgBox("Le client a été modifié, voulez-vous changer les punch de ce projet ?", vbYesNo) = vbYes Then

 Do While Not rstPunch.EOF
 rstPunch.Fields("NoClient") = cmbclient.ItemData(cmbclient.ListIndex)

 Call rstPunch.Update

 Call rstPunch.MoveNext
5  Loop
5  End If
5  End If

5  Call rstPunch.Close
5  Set rstPunch = Nothing
5  End If

5  rstOuvert.Fields("Description") = txtProjet.Text

5  Call rstOuvert.Update

60 Call rstOuvert.Close
  Set rstOuvert = Nothing

 'Si c'est une modification, il faut effacer les pieces et remplir les nouvelles
  Call g_connData.Execute("DELETE * FROM " & sTablePiece & " WHERE " & sChamps & " = '" & sNoProjSoum & "' AND Type = 'E'")

  If m_eType = TYPE_PROJET Then
  If Right$(sNoProjSoum, 2) >= 60 And Right$(sNoProjSoum, 2) <=   Then
  Call g_connData.Execute("DELETE * FROM GrbProjet_Pieces WHERE IDProjet = '" & Left$(sNoProjSoum, 6) & "-" & rstProjSoum.Fields("LiaisonChargeable") & "' AND Type = 'E' AND (PieceExtraChargeable = True OR PieceExtraNonChargeable = True) AND Provenance = '" & Right$(txtNoProjSoum.Text, 2) & "'")
  End If
  End If
  End If
 
 'Enregistrement de la soumission
 'Pour savoir que c'est une soumission ou un projet électrique
  rstProjSoum.Fields("IDClient") = cmbclient.ItemData(cmbclient.ListIndex)
  rstProjSoum.Fields("IDContact") = cmbContact.ItemData(cmbContact.ListIndex)
  rstProjSoum.Fields("description") = txtProjet.Text
6  rstProjSoum.Fields("NbreManuel") = txtNbreManuel.Text
6  rstProjSoum.Fields("transport") = cmbtransport.Text
6  rstProjSoum.Fields("CSA") = chkCSA.Value
6  rstProjSoum.Fields("CUL") = chkCUL.Value
6  rstProjSoum.Fields("UL") = chkUL.Value
6  rstProjSoum.Fields("CUR") = chkCUR.Value
6  rstProjSoum.Fields("UR") = chkUR.Value
6  rstProjSoum.Fields("CE") = chkCE.Value

70 If txtDelais.Text <> "" Then
  rstProjSoum.Fields("Delais") = txtDelais.Text
  Else
  rstProjSoum.Fields("Delais") = " "
  End If

  If m_eType = TYPE_SOUMISSION Then
  rstProjSoum.Fields("TempsDessin") = m_sTempsDessin
  rstProjSoum.Fields("TempsFabrication") = m_sTempsFabrication
  rstProjSoum.Fields("TempsAssemblage") = m_sTempsAssemblage
  rstProjSoum.Fields("TempsProgInterface") = m_sTempsProgInterface
  rstProjSoum.Fields("TempsProgAutomate") = m_sTempsProgAutomate
  rstProjSoum.Fields("TempsProgRobot") = m_sTempsProgRobot
   rstProjSoum.Fields("TempsVision") = m_sTempsVision
   rstProjSoum.Fields("TempsTest") = m_sTempsTest
7  rstProjSoum.Fields("TempsInstallation") = m_sTempsInstallation
7  rstProjSoum.Fields("TempsMiseService") = m_sTempsMiseService
7  rstProjSoum.Fields("TempsFormation") = m_sTempsFormation
7  rstProjSoum.Fields("TempsGestion") = m_sTempsGestion
7  rstProjSoum.Fields("TempsShipping") = m_sTempsShipping
7  End If

80 rstProjSoum.Fields("NbrePersonne") = m_sNbrePersonne
80 rstProjSoum.Fields("TempsHebergement") = m_sTempsHebergement
  rstProjSoum.Fields("TempsRepas") = m_sTempsRepas
  rstProjSoum.Fields("TempsTransport") = m_sTempsTransport
  rstProjSoum.Fields("TempsUniteMobile") = m_sTempsUniteMobile
  rstProjSoum.Fields("PrixEmballage") = m_sPrixEmballage

  rstProjSoum.Fields("TauxHebergement1") = m_sTauxHebergement1
  rstProjSoum.Fields("TauxHebergement2") = m_sTauxHebergement2
  rstProjSoum.Fields("TauxRepas") = m_sTauxRepas
  rstProjSoum.Fields("TauxTransport") = m_sTauxTransport
  rstProjSoum.Fields("TauxUniteMobile") = m_sTauxUniteMobile

  rstProjSoum.Fields("TauxDessin") = m_sTauxDessin
   rstProjSoum.Fields("TauxFabrication") = m_sTauxFabrication
   rstProjSoum.Fields("TauxAssemblage") = m_sTauxAssemblage
   rstProjSoum.Fields("TauxProgInterface") = m_sTauxProgInterface
   rstProjSoum.Fields("TauxProgAutomate") = m_sTauxProgAutomate
8  rstProjSoum.Fields("TauxProgRobot") = m_sTauxProgRobot
8  rstProjSoum.Fields("TauxVision") = m_sTauxVision
8  rstProjSoum.Fields("TauxTest") = m_sTauxTest
8  rstProjSoum.Fields("TauxInstallation") = m_sTauxInstallation
90 rstProjSoum.Fields("TauxMiseService") = m_sTauxMiseService
90 rstProjSoum.Fields("TauxFormation") = m_sTauxFormation
  rstProjSoum.Fields("TauxGestion") = m_sTauxGestion
  rstProjSoum.Fields("TauxShipping") = m_sTauxShipping

  rstProjSoum.Fields("imprevue") = m_sImprevue
  rstProjSoum.Fields("commission") = m_sCommission
  rstProjSoum.Fields("Profit") = m_sProfit
  rstProjSoum.Fields("SansTemps") = m_bSansTemps
  rstProjSoum.Fields("CheminPhotos") = txtCheminPhotos.Text
  rstProjSoum.Fields("MontantForfait") = txtForfait.Text
  rstProjSoum.Fields("InitialeForfait") = Trim$(Replace(lblForfaitInitiale.Caption, "Par :", ""))

  If Not IsNull(rstProjSoum.Fields("NbrePersonne")) Then
 dblNbrePers = CDbl(rstProjSoum.Fields("NbrePersonne"))
   Else
 dblNbrePers = 0
   End If

 If Not IsNull(rstProjSoum.Fields("TempsHebergement")) Then
   dblJoursHebergement = CDbl(rstProjSoum.Fields("TempsHebergement"))
 Else
9  dblJoursHebergement = 0
 End If

100 If Not IsNull(rstProjSoum.Fields("TempsRepas")) Then
1dblJoursRepas = CDbl(rstProjSoum.Fields("TempsRepas"))
10 Else
 dblJoursRepas = 0
10 End If

If Not IsNull(rstProjSoum.Fields("TauxHebergement1")) Then
1dblHebergement1 = CDbl(rstProjSoum.Fields("TauxHebergement1"))
Else
1dblHebergement1 = 0
End If

10 If Not IsNull(rstProjSoum.Fields("TauxHebergement2")) Then
10  dblHebergement2 = CDbl(rstProjSoum.Fields("TauxHebergement2"))
10  Else
10  dblHebergement2 = 0
10  End If

10  If Not IsNull(rstProjSoum.Fields("TauxRepas")) Then
10  dblRepas = CDbl(rstProjSoum.Fields("TauxRepas"))
109Else
10  dblRepas = 0
110End If

110 rstProjSoum.Fields("TotalRepas") = dblNbrePers * dblJoursRepas * dblRepas

11 dblTotalHebergement = 0

11 Do While dblNbrePers > 0
1 If dblNbrePers >= 2 Then
1 dblTotalHebergement = dblTotalHebergement + (dblJoursHebergement * dblHebergement2)

1 dblNbrePers = dblNbrePers - 2
1 Else
1 dblTotalHebergement = dblTotalHebergement + (dblJoursHebergement * dblHebergement1)

1 dblNbrePers = dblNbrePers - 1
1 End If
11 Loop

11  rstProjSoum.Fields("TotalHebergement") = dblTotalHebergement

11  If bAjoutCommande = True Then
 rstProjSoum.Fields("ProchaineCommande") = 1
11  End If

 'Si c'est un projet, il faut enregistrer le prix de réception
1 If m_eType = TYPE_PROJET Then
1 rstProjSoum.Fields("PrixRéception") = txtPrixReception.Text
1 End If

11  If IsNumeric(txtPrixManuel.Text) Then
 rstProjSoum.Fields("Total_Manuel") = txtPrixManuel.Text
1Else
1 rstProjSoum.Fields("Total_Manuel") = "0"
12 End If

12 rstProjSoum.Fields("total_Commission") = txtCommission.Text
12 rstProjSoum.Fields("Total_Profit") = txtProfit.Text
12 rstProjSoum.Fields("PrixTotal") = txtPrixTotal.Text
12 rstProjSoum.Fields("Total_piece") = txtTotalPieces.Text
12 rstProjSoum.Fields("total_imprevue") = txtImprevus.Text
 
12 rstProjSoum.Fields("PrixTotal") = txtPrixTotal.Text
 
12 rstPiece.CursorLocation = adUseServer

12 If m_eType = TYPE_PROJET Then
12  Call rstPiece.Open("SELECT * FROM GrbProjet_Pieces", g_connData, adOpenDynamic, adLockOptimistic)
12  Else
12  Call rstPiece.Open("SELECT * FROM GrbSoumission_Pieces", g_connData, adOpenDynamic, adLockOptimistic)
12  End If

12  If m_eType = TYPE_PROJET Then
1 If g_bModificationFacturation = True Then
1 Call EnregistrerFACT(sNoProjSoum)
1 End If
130End If

 'Enregistrement des pièces
130 For iCompteur = 1 To lvwSoumission.ListItems.count
1 If lvwSoumission.ListItems(iCompteur).Tag <> vbNullString Then
1 If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> vbNullString Then
1 Set itmPiece = lvwSoumission.ListItems(iCompteur)
 
1 Call rstPiece.AddNew

1 If m_eType = TYPE_PROJET Then
1 rstPiece.Fields("IDProjet") = sNoProjSoum
1 Else
1 rstPiece.Fields("IDSoumission") = sNoProjSoum
1 End If
 
1 rstPiece.Fields("Type") = "E"
 
1 If itmPiece.Checked = True Then
1 rstPiece.Fields("Visible") = True
1 Else
1 rstPiece.Fields("Visible") = False
1 End If

1 If m_eType = TYPE_PROJET Then
1 rstPiece.Fields("Facturation") = itmPiece.SubItems(I_COL_SOUM_FACTURATION)

1 If itmPiece.SubItems(I_COL_SOUM_FACTURATION) = "" Then
1 itmPiece.SubItems(I_COL_SOUM_FACTURATION) = " "
14 End If

14 If itmPiece.SubItems(I_COL_SOUM_DATE_COMMANDE) = "" Then
14 itmPiece.SubItems(I_COL_SOUM_DATE_COMMANDE) = " "
14 End If

14 rstPiece.Fields("NoRetour") = itmPiece.ListSubItems(I_COL_SOUM_DATE_COMMANDE).Tag

14 rstPiece.Fields("DateRéception") = itmPiece.ListSubItems(I_COL_SOUM_PRIX_NET).Tag
14 End If
 
14 rstPiece.Fields("IDSection") = itmPiece.Tag
14 rstPiece.Fields("NumItem") = Trim$(itmPiece.SubItems(I_COL_SOUM_PIECE))
14 rstPiece.Fields("Qté") = Replace(itmPiece.Text, "*", vbNullString)

14 If itmPiece.SubItems(I_COL_SOUM_PIECE) = "Texte" Or itmPiece.SubItems(I_COL_SOUM_PIECE) = "Text" Then
14  rstPiece.Fields("DESC_EN") = ""
14  rstPiece.Fields("DESC_FR") = Trim$(itmPiece.SubItems(I_COL_SOUM_DESCR))
14  Else
14  If m_eLangage = ANGLAIS Then
14  rstPiece.Fields("DESC_EN") = Trim$(itmPiece.SubItems(I_COL_SOUM_DESCR))
14  rstPiece.Fields("DESC_FR") = Trim$(itmPiece.ListSubItems(I_COL_SOUM_DESCR).Tag)
14  Else
14  rstPiece.Fields("DESC_FR") = Trim$(itmPiece.SubItems(I_COL_SOUM_DESCR))
150 rstPiece.Fields("DESC_EN") = Trim$(itmPiece.ListSubItems(I_COL_SOUM_DESCR).Tag)
1 End If
 End If

 rstPiece.Fields("Manufact") = Trim$(itmPiece.SubItems(I_COL_SOUM_MANUFACT))
 rstPiece.Fields("Prix_list") = Conversion(itmPiece.SubItems(I_COL_SOUM_PRIX_LIST), MODE_PAS_FORMAT, 4)

 If Trim$(itmPiece.SubItems(I_COL_SOUM_ESCOMPTE)) <> "" Then
 rstPiece.Fields("Escompte") = Conversion(Replace(itmPiece.SubItems(I_COL_SOUM_ESCOMPTE), "%", "") / 100, MODE_PAS_FORMAT)
 Else
 rstPiece.Fields("Escompte") = ""
 End If

 rstPiece.Fields("Prix_net") = Conversion(itmPiece.SubItems(I_COL_SOUM_PRIX_NET), MODE_PAS_FORMAT, 4)

 rstPiece.Fields("OrdreSection") = itmPiece.ListSubItems(I_COL_SOUM_MANUFACT).Tag
15  rstPiece.Fields("NuméroLigne") = iCompteur
 
 'Si le listItem est COLOR_ORANGE
15  If itmPiece.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_ORANGE Then
15  rstPiece.Fields("Commandé") = True
15  Else
15  rstPiece.Fields("Commandé") = False
15  End If

15  If itmPiece.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_GRIS Then
15  rstPiece.Fields("Recu") = True
160 Else
 rstPiece.Fields("Recu") = False
 End If

 If itmPiece.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_ROUGE Then
 rstPiece.Fields("Retour") = True
 Else
 rstPiece.Fields("Retour") = False
 End If

 If itmPiece.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_VERT_FORET And itmPiece.ListSubItems(I_COL_SOUM_PIECE).Bold = True Then
 rstPiece.Fields("CommandeAnnulée") = True
 Else
 rstPiece.Fields("CommandeAnnulée") = False
16  End If

16  If itmPiece.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_BRUN Then
16  rstPiece.Fields("MatérielInutile") = True
16  Else
16  rstPiece.Fields("MatérielInutile") = False
16  End If
 
16  If itmPiece.SubItems(I_COL_SOUM_DISTRIB) <> vbNullString Then
16  rstPiece.Fields("IDFRS") = itmPiece.ListSubItems(I_COL_SOUM_DISTRIB).Tag
170 End If
 
 rstPiece.Fields("Temps") = Trim$(itmPiece.SubItems(I_COL_SOUM_TEMPS))
 rstPiece.Fields("Temps_Total") = itmPiece.SubItems(I_COL_SOUM_MONTAGE)
 rstPiece.Fields("Prix_Total") = Conversion(itmPiece.SubItems(I_COL_SOUM_TOTAL), MODE_PAS_FORMAT)
 rstPiece.Fields("Profit_argent") = Conversion(itmPiece.SubItems(I_COL_SOUM_PROFIT), MODE_PAS_FORMAT)

 If Len(itmPiece.ListSubItems(I_COL_SOUM_PIECE).Tag) <= 50 Then
 rstPiece.Fields("SousSection") = itmPiece.ListSubItems(I_COL_SOUM_PIECE).Tag
 Else
 rstPiece.Fields("SousSection") = Left$(itmPiece.ListSubItems(I_COL_SOUM_PIECE).Tag, 50)
 End If
 
 If itmPiece.SubItems(I_COL_SOUM_PRIX_LIST) <> vbNullString Then
 If itmPiece.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag <> vbNullString Then
1   rstPiece.Fields("PrixOrigine") = Replace(Round(CDbl(Replace(itmPiece.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag, ".", ",")), 4), ".", ",")
1   Else
17  rstPiece.Fields("PrixOrigine") = "0"
17  End If
17  Else
17  rstPiece.Fields("PrixOrigine") = "0"
17  End If

17  If itmPiece.SubItems(I_COL_SOUM_TOTAL) <> "" Then
180 rstPiece.Fields("Devise") = itmPiece.ListSubItems(I_COL_SOUM_TOTAL).Tag
 Else
 rstPiece.Fields("Devise") = ""
 End If
 
 If InStr(1, itmPiece.Text, "*") > 0 Then
 rstPiece.Fields("Quoté") = True
 Else
 rstPiece.Fields("Quoté") = False
 End If

 If m_eType = TYPE_PROJET Then
 If Trim$(itmPiece.SubItems(I_COL_SOUM_ID) <> vbNullString) Then
 rstPiece.Fields("ID") = itmPiece.SubItems(I_COL_SOUM_ID)
1   End If

1   rstPiece.Fields("DateCommande") = itmPiece.SubItems(I_COL_SOUM_DATE_COMMANDE)
1   rstPiece.Fields("DateRequise") = itmPiece.SubItems(I_COL_SOUM_DATE_REQUISE)

1   If itmPiece.SubItems(I_COL_SOUM_DATE_REQUISE) = "" Then
18  itmPiece.SubItems(I_COL_SOUM_DATE_REQUISE) = " "
18  End If

18  rstPiece.Fields("DateRetour") = itmPiece.ListSubItems(I_COL_SOUM_DATE_REQUISE).Tag

18  rstPiece.Fields("NomCommande") = itmPiece.SubItems(I_COL_SOUM_NOM_COMMANDE)

190 rstPiece.Fields("NoSéquentiel") = itmPiece.SubItems(I_COL_SOUM_NO_SEQUENTIEL)
 End If

1  If m_eType = TYPE_PROJET Then
1  If itmPiece.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_ROSE Then
1  rstPiece.Fields("PieceExtraNonChargeable") = True
1  rstPiece.Fields("PieceExtraChargeable") = False
1  Else
1  If itmPiece.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_BLEU Then
1  rstPiece.Fields("PieceExtraChargeable") = True
1  rstPiece.Fields("PieceExtraNonChargeable") = False
1  Else
1  If Left$(itmPiece.ListSubItems(I_COL_SOUM_PROFIT).Tag, 6) = "RETOUR" Or Left$(itmPiece.ListSubItems(I_COL_SOUM_PROFIT).Tag, 10) = "ANNULATION" Then
 sExtra = Right$(itmPiece.ListSubItems(I_COL_SOUM_PROFIT).Tag, 2)

1   If sExtra >= "80" And sExtra <= "98" Then
 rstPiece.Fields("PieceExtraNonChargeable") = True
1   rstPiece.Fields("PieceExtraChargeable") = False
 Else
1   rstPiece.Fields("PieceExtraChargeable") = True
 rstPiece.Fields("PieceExtraNonChargeable") = False
19  End If
200 End If
 End If
 End If

 If itmPiece.SubItems(I_COL_SOUM_PROVENANCE) <> "" Then
 rstPiece.Fields("Provenance") = Right$(itmPiece.SubItems(I_COL_SOUM_PROVENANCE), 2)
 Else
 If Left$(itmPiece.ListSubItems(I_COL_SOUM_PROFIT).Tag, 6) = "RETOUR" Or Left$(itmPiece.ListSubItems(I_COL_SOUM_PROFIT).Tag, 10) = "ANNULATION" Then
 rstPiece.Fields("Provenance") = sExtra
 End If
 End If
 End If

 rstPiece.Fields("Commentaire") = itmPiece.SubItems(I_COL_SOUM_COMMENTAIRE)

20  Call rstPiece.Update

20  If m_eType = TYPE_PROJET Then
20  If Right$(txtNoProjSoum.Text, 2) <=   And Right$(txtNoProjSoum.Text, 2) >= 80 Then
20  Call AjouterPiecesExtraDansJob(itmPiece, rstProjSoum.Fields("LiaisonChargeable"))
20  Else
20  If Right$(txtNoProjSoum.Text, 2) <= 7 And Right$(txtNoProjSoum.Text, 2) >= 60 Then
20  Call AjouterPiecesExtraChargeableDansJob(itmPiece, rstProjSoum.Fields("LiaisonChargeable"))
20  Else
 If Left$(itmPiece.ListSubItems(I_COL_SOUM_PROFIT).Tag, 6) = "RETOUR" Then
 Call AjouterInutileDansExtra(itmPiece, sExtra)

 bCalculExtra = True

 bExiste = False

 For iCompteurExtra = 1 To collExtra.count
 If collExtra(iCompteurExtra) = Right$(itmPiece.ListSubItems(I_COL_SOUM_PROFIT).Tag, 2) Then
 bExiste = True

 Exit For
 End If
 Next

 If bExiste = False Then
 Call collExtra.Add(Right$(itmPiece.ListSubItems(I_COL_SOUM_PROFIT).Tag, 2))
 End If
 Else
 If Left$(itmPiece.ListSubItems(I_COL_SOUM_PROFIT).Tag, 10) = "ANNULATION" Then
 Call AjouterAnnulationDansExtra(itmPiece, sExtra)

 bCalculExtra = True

 bExiste = False
 
 For iCompteurExtra = 1 To collExtra.count
21  If collExtra(iCompteurExtra) = Right$(itmPiece.ListSubItems(I_COL_SOUM_PROFIT).Tag, 2) Then
 bExiste = True

 Exit For
2 End If
2 Next

2 If bExiste = False Then
2 Call collExtra.Add(Right$(itmPiece.ListSubItems(I_COL_SOUM_PROFIT).Tag, 2))
2 End If
2 End If
2 End If
2 End If
2 End If
2 End If
2 End If
2 End If
22  Next

22  If m_eType = TYPE_PROJET Then
22  If Right$(txtNoProjSoum.Text, 2) <=   And Right$(txtNoProjSoum.Text, 2) >= 60 Then
2 Call CalculerTotalRecordset(Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & rstProjSoum.Fields("LiaisonChargeable"))
22  End If
22  End If

230 If bCalculExtra = True Then
23 For iCompteurExtra = 1 To collExtra.count
2 Call CalculerTotalRecordset(Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & collExtra(iCompteurExtra))
2 Next
23 End If
 
23 rstProjSoum.Fields("total_temps") = txtTotalTemps.Text

23 Call rstProjSoum.Update

23 Call rstProjSoum.Close
23 Set rstProjSoum = Nothing

23 Call rstPiece.Close
23 Set rstPiece = Nothing

23 If m_eType = TYPE_SOUMISSION Then
23  Call AjouterSoumissionAuCumulatif
23  Else
23  Call AjouterProjetAuCumulatif
23  End If

238Exit Sub

Oups:

2woups"frmProjSoumElec", "EnregistrerProjSoum", Err, Erl, sNoProjSoum)

 'Si un erreur se produit dans l'enregistrement des pièces, il faut avertir
 'l'utilisateur de quelle pièce il s'agit et continuer avec un Resume Next
23  If Erl >= 1310 And Erl <= 2265 Then
2 If m_eLangage = ANGLAIS Then
2 Call rstSection.Open("SELECT NomSectionEN FROM GrbSoumProjSectionElec", g_connData, adOpenDynamic, adLockOptimistic)
 
24 If Not rstSection.EOF Then
24 sSection = rstSection.Fields("NomSectionEN")
24 Else
24 sSection = itmPiece.Tag
24 End If
24 Else
24 Call rstSection.Open("SELECT NomSectionFR FROM GrbSoumProjSectionElec", g_connData, adOpenDynamic, adLockOptimistic)
 
24 If Not rstSection.EOF Then
24 sSection = rstSection.Fields("NomSectionFR")
24 Else
24 sSection = itmPiece.Tag
24  End If
24  End If
 
24  Call rstSection.Close
24  Set rstSection = Nothing
 
24  Call MsgBox("La pièce " & itmPiece.SubItems(I_COL_SOUM_PIECE) & " de la section " & sSection & " risque de contenir des erreurs." & vbNewLine & _
 "Il se peut qu'elle ne soit plus présente dans la liste.")
24  End If
 
24  Resume Next
End Sub

Private Sub InitialiserNouveauxTaux()

 On Error GoTo Oups

 Dim rstConfig As ADODB.Recordset

 Set rstConfig = New ADODB.Recordset

 Call rstConfig.Open("SELECT TauxDessinElec, TauxFabrication, TauxAssemblageElec, TauxProgInterface, TauxProgAutomate, TauxProgRobot, TauxVision, TauxTestElec, TauxInstallationElec, TauxMiseService, TauxFormationElec, TauxGestionProjetsElec, TauxShippingElec, Hebergement1, Hebergement2, Repas, Standard, UniteMobile FROM GrbConfig", g_connData, adOpenDynamic, adLockOptimistic)

 If Not IsNull(rstConfig.Fields("TauxDessinElec")) Then
 m_sTauxDessin = rstConfig.Fields("TauxDessinElec")
 Else
 m_sTauxDessin = "0"
 End If

 If Not IsNull(rstConfig.Fields("TauxFabrication")) Then
 m_sTauxFabrication = rstConfig.Fields("TauxFabrication")
  Else
  m_sTauxFabrication = "0"
  End If

  If Not IsNull(rstConfig.Fields("TauxAssemblageElec")) Then
  m_sTauxAssemblage = rstConfig.Fields("TauxAssemblageElec")
  Else
  m_sTauxAssemblage = "0"
  End If

10 If Not IsNull(rstConfig.Fields("TauxProgInterface")) Then
1 m_sTauxProgInterface = rstConfig.Fields("TauxProgInterface")
Else
 m_sTauxProgInterface = "0"
End If

If Not IsNull(rstConfig.Fields("TauxProgAutomate")) Then
 m_sTauxProgAutomate = rstConfig.Fields("TauxProgAutomate")
Else
 m_sTauxProgAutomate = "0"
End If

If Not IsNull(rstConfig.Fields("TauxProgRobot")) Then
 m_sTauxProgRobot = rstConfig.Fields("TauxProgRobot")
1  Else
 m_sTauxProgRobot = "0"
 End If

If Not IsNull(rstConfig.Fields("TauxVision")) Then
 m_sTauxVision = rstConfig.Fields("TauxVision")
Else
 m_sTauxVision = "0"
1  End If

 If Not IsNull(rstConfig.Fields("TauxTestElec")) Then
 m_sTauxTest = rstConfig.Fields("TauxTestElec")
Else
 m_sTauxTest = "0"
End If

If Not IsNull(rstConfig.Fields("TauxInstallationElec")) Then
 m_sTauxInstallation = rstConfig.Fields("TauxInstallationElec")
Else
 m_sTauxInstallation = "0"
End If

If Not IsNull(rstConfig.Fields("TauxMiseService")) Then
 m_sTauxMiseService = rstConfig.Fields("TauxMiseService")
2  Else
 m_sTauxMiseService = "0"
2  End If

If Not IsNull(rstConfig.Fields("TauxFormationElec")) Then
m_sTauxFormation = rstConfig.Fields("TauxFormationElec")
Else
m_sTauxFormation = "0"
End If

30 If Not IsNull(rstConfig.Fields("TauxGestionProjetsElec")) Then
3 m_sTauxGestion = rstConfig.Fields("TauxGestionProjetsElec")
Else
 m_sTauxGestion = "0"
End If

If Not IsNull(rstConfig.Fields("TauxShippingElec")) Then
 m_sTauxShipping = rstConfig.Fields("TauxShippingElec")
Else
 m_sTauxShipping = "0"
End If

If m_eType = TYPE_PROJET Then
 m_sTauxHebergement1 = "0"
m_sTauxHebergement2 = "0"
 m_sTauxRepas = "0"
m_sTauxTransport = "0"
 m_sTauxUniteMobile = "0"
3  Else
 m_sTauxHebergement1 = rstConfig.Fields("Hebergement1")
 m_sTauxHebergement2 = rstConfig.Fields("Hebergement2")
 m_sTauxRepas = rstConfig.Fields("Repas")
m_sTauxTransport = rstConfig.Fields("Standard")
4 m_sTauxUniteMobile = rstConfig.Fields("UniteMobile")
4 End If

4 Call rstConfig.Close
4 Set rstConfig = Nothing

4 Exit Sub

Oups:

4 wOups "frmProjSoumElec", "InitialiserNouveauxTaux", Err, Err.number, Err.Description
End Sub

Private Sub AjouterPiecesExtraChargeableDansJob(ByVal itmSource As ListItem, ByVal sLiaison As String)

 On Error GoTo Oups

 Dim rstPiece As ADODB.Recordset
 Dim rstProjet As ADODB.Recordset
 Dim rstSection As ADODB.Recordset
 Dim iCompteur As Integer
 Dim sSection As String
 Dim bSkip As Boolean
 
 Set rstProjet = New ADODB.Recordset
 
 Call rstProjet.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sLiaison & "'", g_connData, adOpenDynamic, adLockOptimistic)

 'Si le projet existe
 If Not rstProjet.EOF Then
 'Ouverture du recordset sur le projet original
  Set rstPiece = New ADODB.Recordset
 
  Call rstPiece.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sLiaison & "' AND IDSection = " & itmSource.Tag & " AND SousSection = '" & Replace(itmSource.ListSubItems(I_COL_SOUM_PIECE).Tag, "'", "''") & "' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)

  If Not rstPiece.EOF Then
  Call rstPiece.MoveLast

  iCompteur = rstPiece.Fields("NuméroLigne") + 1
  Else
  iCompteur = 1
  End If

Call rstPiece.AddNew

1 rstPiece.Fields("IDProjet") = Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sLiaison
 
 rstPiece.Fields("Type") = "E"
 
 If itmSource.Checked = True Then
 rstPiece.Fields("Visible") = True
 Else
 rstPiece.Fields("Visible") = False
 End If

 rstPiece.Fields("Facturation") = itmSource.SubItems(I_COL_SOUM_FACTURATION)
 
 rstPiece.Fields("IDSection") = itmSource.Tag
 rstPiece.Fields("NumItem") = Trim$(itmSource.SubItems(I_COL_SOUM_PIECE))
 rstPiece.Fields("Qté") = Replace(itmSource.Text, "*", vbNullString)
rstPiece.Fields("Desc_FR") = Trim$(itmSource.SubItems(I_COL_SOUM_DESCR))
 rstPiece.Fields("Desc_EN") = Trim$(itmSource.ListSubItems(I_COL_SOUM_DESCR).Tag)
 rstPiece.Fields("Manufact") = Trim$(itmSource.SubItems(I_COL_SOUM_MANUFACT))
 rstPiece.Fields("Prix_list") = Conversion(itmSource.SubItems(I_COL_SOUM_PRIX_LIST), MODE_PAS_FORMAT, 4)
 rstPiece.Fields("Escompte") = Conversion(itmSource.SubItems(I_COL_SOUM_ESCOMPTE), MODE_PAS_FORMAT)
 rstPiece.Fields("Prix_net") = Conversion(itmSource.SubItems(I_COL_SOUM_PRIX_NET), MODE_PAS_FORMAT, 4)
 rstPiece.Fields("OrdreSection") = itmSource.ListSubItems(I_COL_SOUM_MANUFACT).Tag
1  rstPiece.Fields("NuméroLigne") = iCompteur
 
 If itmSource.SubItems(I_COL_SOUM_DISTRIB) <> vbNullString Then
 rstPiece.Fields("IDFRS") = itmSource.ListSubItems(I_COL_SOUM_DISTRIB).Tag
 End If
 
 rstPiece.Fields("Temps") = Trim$(itmSource.SubItems(I_COL_SOUM_TEMPS))
 rstPiece.Fields("Temps_Total") = itmSource.SubItems(I_COL_SOUM_MONTAGE)
 rstPiece.Fields("Prix_Total") = Conversion(itmSource.SubItems(I_COL_SOUM_TOTAL), MODE_PAS_FORMAT)
 rstPiece.Fields("Profit_argent") = Conversion(itmSource.SubItems(I_COL_SOUM_PROFIT), MODE_PAS_FORMAT)
 rstPiece.Fields("SousSection") = itmSource.ListSubItems(I_COL_SOUM_PIECE).Tag
 
 If itmSource.SubItems(I_COL_SOUM_PRIX_LIST) <> vbNullString Then
 If itmSource.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag <> vbNullString Then
 rstPiece.Fields("PrixOrigine") = Replace(Round(CDbl(Replace(itmSource.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag, ".", ",")), 4), ".", ",")
 Else
 rstPiece.Fields("PrixOrigine") = "0"
 End If
Else
 rstPiece.Fields("PrixOrigine") = "0"
End If
 
 If InStr(1, itmSource.Text, "*") > 0 Then
 rstPiece.Fields("Quoté") = True
 Else
 rstPiece.Fields("Quoté") = False
3 End If

 If Trim$(itmSource.SubItems(I_COL_SOUM_ID) <> vbNullString) Then
 rstPiece.Fields("ID") = itmSource.SubItems(I_COL_SOUM_ID)
 End If

 rstPiece.Fields("PieceExtraChargeable") = True
 rstPiece.Fields("Provenance") = Right$(txtNoProjSoum.Text, 2)

 Call rstPiece.Update

 Call rstPiece.Close

 rstPiece.CursorLocation = adUseServer

 Call rstPiece.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sLiaison & "' AND NuméroLigne >= " & iCompteur & " ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)

 'Tant qu'il y a des enregistrements dans le recordset
 Do While Not rstPiece.EOF
 If rstPiece.Fields("PieceExtraChargeable") = True And rstPiece.Fields("Qté") = itmSource.Text And rstPiece.Fields("NumItem") = itmSource.SubItems(I_COL_SOUM_PIECE) And bSkip = False Then
 bSkip = True
 Else
 rstPiece.Fields("NuméroLigne") = rstPiece.Fields("NuméroLigne") + 1

 Call rstPiece.Update
 End If

 Call rstPiece.MoveNext
 Loop

Call rstPiece.Close
4 Set rstPiece = Nothing

4 Call CalculerTempsFabricationRecordset(Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sLiaison)
4 End If

4 Call rstProjet.Close
4 Set rstProjet = Nothing

4 Exit Sub

Oups:

4 wOups "frmProjSoumElec", "AjouterPiecesExtraDansJob", Err, Err.number, Err.Description

 Set rstSection = New ADODB.Recordset
 
 If m_eLangage = ANGLAIS Then
 Call rstSection.Open("SELECT NomSectionEN FROM GrbSoumProjSectionElec", g_connData, adOpenDynamic, adLockOptimistic)
 
 If Not rstSection.EOF Then
 sSection = rstSection.Fields("NomSectionEN")
 Else
 sSection = itmSource.Tag
 End If
 Else
 Call rstSection.Open("SELECT NomSectionFR FROM GrbSoumProjSectionElec", g_connData, adOpenDynamic, adLockOptimistic)
 
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
 
 On Error GoTo Oups

 Dim rstPiece As ADODB.Recordset
 Dim rstProjet As ADODB.Recordset
 Dim rstSection As ADODB.Recordset
 Dim iCompteur As Integer
 Dim sSection As String
 Dim bSkip As Boolean

 Set rstProjet = New ADODB.Recordset
 
 Call rstProjet.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sLiaison & "'", g_connData, adOpenDynamic, adLockOptimistic)

 'Si le projet existe
 If Not rstProjet.EOF Then
 'Ouverture du recordset sur le projet original
 Set rstPiece = New ADODB.Recordset
 
  Call rstPiece.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sLiaison & "' AND IDSection = " & itmSource.Tag & " AND SousSection = '" & Replace(itmSource.ListSubItems(I_COL_SOUM_PIECE).Tag, "'", "''") & "' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)

  If Not rstPiece.EOF Then
  Call rstPiece.MoveLast

  iCompteur = rstPiece.Fields("NuméroLigne") + 1
  Else
  iCompteur = 1
  End If

  Call rstPiece.AddNew

rstPiece.Fields("IDProjet") = Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sLiaison
 
1 rstPiece.Fields("Type") = "E"
 
 If itmSource.Checked = True Then
 rstPiece.Fields("Visible") = True
 Else
 rstPiece.Fields("Visible") = False
 End If

 rstPiece.Fields("Facturation") = itmSource.SubItems(I_COL_SOUM_FACTURATION)
 
 rstPiece.Fields("IDSection") = itmSource.Tag
 rstPiece.Fields("NumItem") = Trim$(itmSource.SubItems(I_COL_SOUM_PIECE))
 rstPiece.Fields("Qté") = Replace(itmSource.Text, "*", vbNullString)
 rstPiece.Fields("Desc_FR") = Trim$(itmSource.SubItems(I_COL_SOUM_DESCR))
rstPiece.Fields("Desc_EN") = Trim$(itmSource.ListSubItems(I_COL_SOUM_DESCR).Tag)
 rstPiece.Fields("Manufact") = Trim$(itmSource.SubItems(I_COL_SOUM_MANUFACT))
 rstPiece.Fields("Prix_list") = Conversion(itmSource.SubItems(I_COL_SOUM_PRIX_LIST), MODE_PAS_FORMAT, 4)
 rstPiece.Fields("Escompte") = Conversion(itmSource.SubItems(I_COL_SOUM_ESCOMPTE), MODE_PAS_FORMAT)
 rstPiece.Fields("Prix_net") = Conversion(itmSource.SubItems(I_COL_SOUM_PRIX_NET), MODE_PAS_FORMAT, 4)
 rstPiece.Fields("OrdreSection") = itmSource.ListSubItems(I_COL_SOUM_MANUFACT).Tag
 rstPiece.Fields("NuméroLigne") = iCompteur
 
1  If itmSource.SubItems(I_COL_SOUM_DISTRIB) <> vbNullString Then
 rstPiece.Fields("IDFRS") = itmSource.ListSubItems(I_COL_SOUM_DISTRIB).Tag
 End If
 
 rstPiece.Fields("Temps") = Trim$(itmSource.SubItems(I_COL_SOUM_TEMPS))
 rstPiece.Fields("Temps_Total") = itmSource.SubItems(I_COL_SOUM_MONTAGE)
 rstPiece.Fields("Prix_Total") = Conversion(itmSource.SubItems(I_COL_SOUM_TOTAL), MODE_PAS_FORMAT)
 rstPiece.Fields("Profit_argent") = Conversion(itmSource.SubItems(I_COL_SOUM_PROFIT), MODE_PAS_FORMAT)
 rstPiece.Fields("SousSection") = itmSource.ListSubItems(I_COL_SOUM_PIECE).Tag
 
 If itmSource.SubItems(I_COL_SOUM_PRIX_LIST) <> vbNullString Then
 If itmSource.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag <> vbNullString Then
 rstPiece.Fields("PrixOrigine") = Replace(Round(CDbl(Replace(itmSource.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag, ".", ",")), 2), ".", ",")
 Else
 rstPiece.Fields("PrixOrigine") = "0"
 End If
 Else
 rstPiece.Fields("PrixOrigine") = "0"
 End If
 
If InStr(1, itmSource.Text, "*") > 0 Then
 rstPiece.Fields("Quoté") = True
Else
 rstPiece.Fields("Quoté") = False
End If

3 If Trim$(itmSource.SubItems(I_COL_SOUM_ID) <> vbNullString) Then
 rstPiece.Fields("ID") = itmSource.SubItems(I_COL_SOUM_ID)
 End If

 rstPiece.Fields("PieceExtraNonChargeable") = True
 rstPiece.Fields("Provenance") = Right$(txtNoProjSoum.Text, 2)

 Call rstPiece.Update

 Call rstPiece.Close

 rstPiece.CursorLocation = adUseServer

 Call rstPiece.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sLiaison & "' AND NuméroLigne >= " & iCompteur & " ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)

 'Tant qu'il y a des enregistrements dans le recordset
 Do While Not rstPiece.EOF
 If rstPiece.Fields("PieceExtraNonChargeable") = True And rstPiece.Fields("Qté") = itmSource.Text And rstPiece.Fields("NumItem") = itmSource.SubItems(I_COL_SOUM_PIECE) And bSkip = False Then
 bSkip = True
 Else
 rstPiece.Fields("NuméroLigne") = rstPiece.Fields("NuméroLigne") + 1

 Call rstPiece.Update
 End If

 Call rstPiece.MoveNext
 Loop

 Call rstPiece.Close
Set rstPiece = Nothing

4 Call CalculerTempsFabricationRecordset(Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sLiaison)
4 End If

4 Call rstProjet.Close
4 Set rstProjet = Nothing

4 Exit Sub

Oups:

4 wOups "frmProjSoumElec", "AjouterPiecesExtraDansJob", Err, Err.number, Err.Description

 Set rstSection = New ADODB.Recordset

 If m_eLangage = ANGLAIS Then
 Call rstSection.Open("SELECT NomSectionEN FROM GrbSoumProjSectionElec", g_connData, adOpenDynamic, adLockOptimistic)
 
 If Not rstSection.EOF Then
 sSection = rstSection.Fields("NomSectionEN")
 Else
 sSection = itmSource.Tag
 End If
 Else
 Call rstSection.Open("SELECT NomSectionFR FROM GrbSoumProjSectionElec", g_connData, adOpenDynamic, adLockOptimistic)
 
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
 
 On Error GoTo Oups

 Dim rstPiece As ADODB.Recordset
 Dim rstProjet As ADODB.Recordset
 Dim rstSection As ADODB.Recordset
 Dim iCompteur As Integer
 Dim sSection As String
 Dim bSkip As Boolean
 
 Set rstProjet = New ADODB.Recordset
 
 Call rstProjet.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sExtra & "'", g_connData, adOpenDynamic, adLockOptimistic)

 'Si le projet existe
 If Not rstProjet.EOF Then
 'Ouverture du recordset sur le projet original
 Set rstPiece = New ADODB.Recordset
 
  Call rstPiece.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sExtra & "' AND IDSection = " & itmSource.Tag & " AND SousSection = '" & Replace(itmSource.ListSubItems(I_COL_SOUM_PIECE).Tag, "'", "''") & "' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)

  If Not rstPiece.EOF Then
  Call rstPiece.MoveLast

  iCompteur = rstPiece.Fields("NuméroLigne") + 1
  Else
  iCompteur = 1
  End If

  Call rstPiece.AddNew

rstPiece.Fields("IDProjet") = Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sExtra
 
1 rstPiece.Fields("Type") = "E"
 
 If itmSource.Checked = True Then
 rstPiece.Fields("Visible") = True
 Else
 rstPiece.Fields("Visible") = False
 End If
 
 rstPiece.Fields("IDSection") = itmSource.Tag
 rstPiece.Fields("NumItem") = Trim$(itmSource.SubItems(I_COL_SOUM_PIECE))
 rstPiece.Fields("Qté") = Replace(itmSource.Text, "*", vbNullString)

 If m_eLangage = ANGLAIS Then
 rstPiece.Fields("DESC_EN") = Trim$(itmSource.SubItems(I_COL_SOUM_DESCR))
 rstPiece.Fields("DESC_FR") = Trim$(itmSource.ListSubItems(I_COL_SOUM_DESCR).Tag)
 Else
 rstPiece.Fields("DESC_FR") = Trim$(itmSource.SubItems(I_COL_SOUM_DESCR))
 rstPiece.Fields("DESC_EN") = Trim$(itmSource.ListSubItems(I_COL_SOUM_DESCR).Tag)
 End If

 rstPiece.Fields("Manufact") = Trim$(itmSource.SubItems(I_COL_SOUM_MANUFACT))
 rstPiece.Fields("Prix_list") = Conversion(itmSource.SubItems(I_COL_SOUM_PRIX_LIST), MODE_PAS_FORMAT, 4)
1  rstPiece.Fields("Escompte") = Conversion(itmSource.SubItems(I_COL_SOUM_ESCOMPTE), MODE_PAS_FORMAT)
 rstPiece.Fields("Prix_net") = Conversion(itmSource.SubItems(I_COL_SOUM_PRIX_NET), MODE_PAS_FORMAT, 4)

 rstPiece.Fields("OrdreSection") = itmSource.ListSubItems(I_COL_SOUM_MANUFACT).Tag
 rstPiece.Fields("NuméroLigne") = iCompteur
 
 If itmSource.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_NOIR Then
 rstPiece.Fields("MatérielInutile") = False
 Else
 rstPiece.Fields("MatérielInutile") = True
 End If
 
 If itmSource.SubItems(I_COL_SOUM_DISTRIB) <> vbNullString Then
 rstPiece.Fields("IDFRS") = itmSource.ListSubItems(I_COL_SOUM_DISTRIB).Tag
 End If
 
 rstPiece.Fields("Temps") = Trim$(itmSource.SubItems(I_COL_SOUM_TEMPS))
rstPiece.Fields("Temps_Total") = itmSource.SubItems(I_COL_SOUM_MONTAGE)
 rstPiece.Fields("Prix_Total") = Conversion(itmSource.SubItems(I_COL_SOUM_TOTAL), MODE_PAS_FORMAT)
rstPiece.Fields("Profit_argent") = Conversion(itmSource.SubItems(I_COL_SOUM_PROFIT), MODE_PAS_FORMAT)
 rstPiece.Fields("SousSection") = itmSource.ListSubItems(I_COL_SOUM_PIECE).Tag
 
If itmSource.SubItems(I_COL_SOUM_PRIX_LIST) <> vbNullString Then
 If itmSource.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag <> vbNullString Then
 rstPiece.Fields("PrixOrigine") = Replace(Round(CDbl(Replace(itmSource.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag, ".", ",")), 4), ".", ",")
 Else
 rstPiece.Fields("PrixOrigine") = "0"
End If
 Else
 rstPiece.Fields("PrixOrigine") = "0"
 End If
 
 If InStr(1, itmSource.Text, "*") > 0 Then
 rstPiece.Fields("Quoté") = True
 Else
 rstPiece.Fields("Quoté") = False
 End If

 rstPiece.Fields("DateRetour") = itmSource.ListSubItems(I_COL_SOUM_DATE_REQUISE).Tag

 rstPiece.Fields("Commentaire") = itmSource.SubItems(I_COL_SOUM_COMMENTAIRE)

Call rstPiece.Update

 Call rstPiece.Close

rstPiece.CursorLocation = adUseServer

 Call rstPiece.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sExtra & "' AND NuméroLigne >= " & iCompteur & " ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)

 'Tant qu'il y a des enregistrements dans le recordset
Do While Not rstPiece.EOF
 If itmSource.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_BRUN Then
 If rstPiece.Fields("MatérielInutile") = True And rstPiece.Fields("Qté") = itmSource.Text And rstPiece.Fields("NumItem") = itmSource.SubItems(I_COL_SOUM_PIECE) And bSkip = False Then
 bSkip = True
 Else
4 rstPiece.Fields("NuméroLigne") = rstPiece.Fields("NuméroLigne") + 1

4 Call rstPiece.Update
4 End If
4 Else
4 If rstPiece.Fields("MatérielInutile") = False And rstPiece.Fields("Qté") = itmSource.Text And rstPiece.Fields("NumItem") = itmSource.SubItems(I_COL_SOUM_PIECE) And bSkip = False Then
4 bSkip = True
4 Else
4 rstPiece.Fields("NuméroLigne") = rstPiece.Fields("NuméroLigne") + 1

4 Call rstPiece.Update
4 End If
4 End If

4  Call rstPiece.MoveNext
4  Loop

4  Call rstPiece.Close
4  Set rstPiece = Nothing

4  Call CalculerTempsFabricationRecordset(Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sExtra)
4  End If

4  Call rstProjet.Close
4  Set rstProjet = Nothing

50 Exit Sub

Oups:

50 wOups "frmProjSoumElec", "AjouterInutileDansExtra", Err, Err.number, Err.Description

 Set rstSection = New ADODB.Recordset

 If m_eLangage = ANGLAIS Then
 Call rstSection.Open("SELECT NomSectionEN FROM GrbSoumProjSectionElec", g_connData, adOpenDynamic, adLockOptimistic)
 
 If Not rstSection.EOF Then
 sSection = rstSection.Fields("NomSectionEN")
 Else
 sSection = itmSource.Tag
 End If
 Else
 Call rstSection.Open("SELECT NomSectionFR FROM GrbSoumProjSectionElec", g_connData, adOpenDynamic, adLockOptimistic)
 
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
 
 On Error GoTo Oups

 Dim rstPiece As ADODB.Recordset
 Dim rstProjet As ADODB.Recordset
 Dim rstSection As ADODB.Recordset
 Dim iCompteur As Integer
 Dim sSection As String
 Dim bSkip As Boolean
 
 Set rstProjet = New ADODB.Recordset
 
 Call rstProjet.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sExtra & "'", g_connData, adOpenDynamic, adLockOptimistic)

 'Si le projet existe
 If Not rstProjet.EOF Then
 'Ouverture du recordset sur le projet original
 Set rstPiece = New ADODB.Recordset
 
  Call rstPiece.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sExtra & "' AND IDSection = " & itmSource.Tag & " AND SousSection = '" & Replace(itmSource.ListSubItems(I_COL_SOUM_PIECE).Tag, "'", "''") & "' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)

  If Not rstPiece.EOF Then
  Call rstPiece.MoveLast

  iCompteur = rstPiece.Fields("NuméroLigne") + 1
  Else
  iCompteur = 1
  End If

  Call rstPiece.AddNew

rstPiece.Fields("IDProjet") = Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sExtra
 
1 rstPiece.Fields("Type") = "E"
 
 If itmSource.Checked = True Then
 rstPiece.Fields("Visible") = True
 Else
 rstPiece.Fields("Visible") = False
 End If
 
 rstPiece.Fields("IDSection") = itmSource.Tag
 rstPiece.Fields("NumItem") = Trim$(itmSource.SubItems(I_COL_SOUM_PIECE))
 rstPiece.Fields("Qté") = Replace(itmSource.Text, "*", vbNullString)

 If m_eLangage = ANGLAIS Then
 rstPiece.Fields("DESC_EN") = Trim$(itmSource.SubItems(I_COL_SOUM_DESCR))
 rstPiece.Fields("DESC_FR") = Trim$(itmSource.ListSubItems(I_COL_SOUM_DESCR).Tag)
 Else
 rstPiece.Fields("DESC_FR") = Trim$(itmSource.SubItems(I_COL_SOUM_DESCR))
 rstPiece.Fields("DESC_EN") = Trim$(itmSource.ListSubItems(I_COL_SOUM_DESCR).Tag)
 End If

 rstPiece.Fields("Manufact") = Trim$(itmSource.SubItems(I_COL_SOUM_MANUFACT))
 rstPiece.Fields("Prix_list") = Conversion(itmSource.SubItems(I_COL_SOUM_PRIX_LIST), MODE_PAS_FORMAT, 4)
1  rstPiece.Fields("Escompte") = Conversion(itmSource.SubItems(I_COL_SOUM_ESCOMPTE), MODE_PAS_FORMAT)
 rstPiece.Fields("Prix_net") = Conversion(itmSource.SubItems(I_COL_SOUM_PRIX_NET), MODE_PAS_FORMAT, 4)

 rstPiece.Fields("OrdreSection") = itmSource.ListSubItems(I_COL_SOUM_MANUFACT).Tag
 rstPiece.Fields("NuméroLigne") = iCompteur
 
 rstPiece.Fields("CommandeAnnulée") = True
 
 If itmSource.SubItems(I_COL_SOUM_DISTRIB) <> vbNullString Then
 rstPiece.Fields("IDFRS") = itmSource.ListSubItems(I_COL_SOUM_DISTRIB).Tag
 End If
 
 rstPiece.Fields("Temps") = Trim$(itmSource.SubItems(I_COL_SOUM_TEMPS))
 rstPiece.Fields("Temps_Total") = itmSource.SubItems(I_COL_SOUM_MONTAGE)
 rstPiece.Fields("Prix_Total") = Conversion(itmSource.SubItems(I_COL_SOUM_TOTAL), MODE_PAS_FORMAT)
 rstPiece.Fields("Profit_argent") = Conversion(itmSource.SubItems(I_COL_SOUM_PROFIT), MODE_PAS_FORMAT)
 rstPiece.Fields("SousSection") = itmSource.ListSubItems(I_COL_SOUM_PIECE).Tag
 
If itmSource.SubItems(I_COL_SOUM_PRIX_LIST) <> vbNullString Then
 If itmSource.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag <> vbNullString Then
 rstPiece.Fields("PrixOrigine") = Replace(Round(CDbl(Replace(itmSource.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag, ".", ",")), 2), ".", ",")
 Else
 rstPiece.Fields("PrixOrigine") = "0"
 End If
Else
 rstPiece.Fields("PrixOrigine") = "0"
End If
 
3 If InStr(1, itmSource.Text, "*") > 0 Then
 rstPiece.Fields("Quoté") = True
 Else
 rstPiece.Fields("Quoté") = False
 End If

 rstPiece.Fields("Commentaire") = itmSource.SubItems(I_COL_SOUM_COMMENTAIRE)

 Call rstPiece.Update

 Call rstPiece.Close

 rstPiece.CursorLocation = adUseServer

 Call rstPiece.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sExtra & "' AND NuméroLigne >= " & iCompteur & " ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)

 'Tant qu'il y a des enregistrements dans le recordset
 Do While Not rstPiece.EOF
 If rstPiece.Fields("CommandeAnnulée") = True And rstPiece.Fields("Qté") = itmSource.Text And rstPiece.Fields("NumItem") = itmSource.SubItems(I_COL_SOUM_PIECE) And bSkip = False Then
 bSkip = True
 Else
 rstPiece.Fields("NuméroLigne") = rstPiece.Fields("NuméroLigne") + 1

 Call rstPiece.Update
 End If

 Call rstPiece.MoveNext
 Loop

Call rstPiece.Close
4 Set rstPiece = Nothing

4 Call CalculerTempsFabricationRecordset(Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sExtra)
4 End If

4 Call rstProjet.Close
4 Set rstProjet = Nothing

4 Exit Sub

Oups:

4 wOups "frmProjSoumElec", "AjouterAnnulationDansExtra", Err, Err.number, Err.Description

 Set rstSection = New ADODB.Recordset

 If m_eLangage = ANGLAIS Then
 Call rstSection.Open("SELECT NomSectionEN FROM GrbSoumProjSectionElec", g_connData, adOpenDynamic, adLockOptimistic)
 
 If Not rstSection.EOF Then
 sSection = rstSection.Fields("NomSectionEN")
 Else
 sSection = itmSource.Tag
 End If
 Else
 Call rstSection.Open("SELECT NomSectionFR FROM GrbSoumProjSectionElec", g_connData, adOpenDynamic, adLockOptimistic)
 
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

 On Error GoTo Oups

 'Fermeture de la fenêtre
 m_bResize = False
 
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "cmdFermer_Click", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListViewModifications()

 On Error GoTo Oups

 'Rempli le lvwHistorique
 Dim rstProjSoum As ADODB.Recordset
 Dim rstEmploye As ADODB.Recordset
 Dim rstCreation As ADODB.Recordset
 Dim sChamps As String
 Dim sTable As String
 Dim sTableCreer As String
 Dim itmModif As ListItem
 
 Set rstProjSoum = New ADODB.Recordset
 Set rstEmploye = New ADODB.Recordset
 Set rstCreation = New ADODB.Recordset
 
 'Il faut le vider avant de le remplir
  Call lvwHistorique.ListItems.Clear
 
  If m_eType = TYPE_PROJET Then
  sChamps = "IDProjet"
  sTable = "GrbProjet_Modif"
  sTableCreer = "GrbProjetElec"
  Else
  sChamps = "IDSoumission"
  sTable = "GrbSoumission_Modif"
sTableCreer = "GrbSoumissionElec"
End If
 
 'Ouverture du recordset selon le type
Call rstCreation.Open("SELECT creer, creer_par FROM " & sTableCreer & " WHERE " & sChamps & " = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Ajout de la section "Création"
Set itmModif = lvwHistorique.ListItems.Add
 
itmModif.Text = "Création"
 
itmModif.Bold = True
 
 'Ajout du nom de celui qui l'a créé
Set itmModif = lvwHistorique.ListItems.Add
 
Call rstEmploye.Open("SELECT Employe FROM GrbEmployés WHERE noEmploye = " & rstCreation.Fields("creer_par"), g_connData, adOpenDynamic, adLockOptimistic)
 
itmModif.Text = rstEmploye.Fields("Employe")
 
Call rstEmploye.Close
 
 'Date
itmModif.SubItems(I_COL_MODIF_DATE) = rstCreation.Fields("creer")
 
itmModif.SubItems(I_COL_MODIF_HEURE) = vbNullString
 
1  Call rstCreation.Close
Set rstCreation = Nothing
 
 Call rstProjSoum.Open("SELECT * FROM " & sTable & " WHERE " & sChamps & " = '" & txtNoProjSoum.Text & "' AND Type = 'E' AND TypeModif = 'MODIFICATION' ORDER BY [Date], Heure", g_connData, adOpenDynamic, adLockOptimistic)
 
If Not rstProjSoum.EOF Then
 'Ajout de la section "Modifications"
 Set itmModif = lvwHistorique.ListItems.Add
 
 itmModif.Text = "Modifications"
 
 itmModif.Bold = True
 
1  Do While Not rstProjSoum.EOF
 'Ajout des modifications
 Set itmModif = lvwHistorique.ListItems.Add
 
 Call rstEmploye.Open("SELECT Employe FROM GrbEmployés WHERE noEmploye = " & rstProjSoum.Fields("NoEmployé"), g_connData, adOpenDynamic, adLockOptimistic)
 
 'Employé
 itmModif.Text = rstEmploye.Fields("Employe")
 
 Call rstEmploye.Close
 
 'Date
 itmModif.SubItems(I_COL_MODIF_DATE) = rstProjSoum.Fields("Date")
 
 'Heure
 itmModif.SubItems(I_COL_MODIF_HEURE) = rstProjSoum.Fields("Heure")
 
 Call rstProjSoum.MoveNext
 Loop
End If
 
Call rstProjSoum.Close

Call rstProjSoum.Open("SELECT * FROM " & sTable & " WHERE " & sChamps & " = '" & txtNoProjSoum.Text & "' AND Type = 'E' AND TypeModif = 'RECEPTION' ORDER BY [Date], Heure", g_connData, adOpenDynamic, adLockOptimistic)
 
If Not rstProjSoum.EOF Then
 'Ajout de la section "Réception"
Set itmModif = lvwHistorique.ListItems.Add
 
 itmModif.Text = "Réception"
 
itmModif.Bold = True
 
 Do While Not rstProjSoum.EOF
 'Ajout des modifications
 Set itmModif = lvwHistorique.ListItems.Add
 
 Call rstEmploye.Open("SELECT Employe FROM GrbEmployés WHERE noEmploye = " & rstProjSoum.Fields("NoEmployé"), g_connData, adOpenDynamic, adLockOptimistic)
 
 'Employé
 itmModif.Text = rstEmploye.Fields("Employe")
 
 Call rstEmploye.Close
 
 'Date
 itmModif.SubItems(I_COL_MODIF_DATE) = rstProjSoum.Fields("Date")
 
 'Heure
itmModif.SubItems(I_COL_MODIF_HEURE) = rstProjSoum.Fields("Heure")
 
 Call rstProjSoum.MoveNext
 Loop
End If
 
Call rstProjSoum.Close

Call rstProjSoum.Open("SELECT * FROM " & sTable & " WHERE " & sChamps & " = '" & txtNoProjSoum.Text & "' AND Type = 'E' AND TypeModif = 'RETOUR' ORDER BY [Date], Heure", g_connData, adOpenDynamic, adLockOptimistic)
 
If Not rstProjSoum.EOF Then
 'Ajout de la section "Retour de marchandise"
 Set itmModif = lvwHistorique.ListItems.Add
 
 itmModif.Text = "Retour de marchandise"
 
 itmModif.Bold = True
 
 Do While Not rstProjSoum.EOF
 'Ajout des retours de marchandise
 Set itmModif = lvwHistorique.ListItems.Add
 
 Call rstEmploye.Open("SELECT Employe FROM GrbEmployés WHERE noEmploye = " & rstProjSoum.Fields("NoEmployé"), g_connData, adOpenDynamic, adLockOptimistic)
 
 'Employé
 itmModif.Text = rstEmploye.Fields("Employe")
 
 Call rstEmploye.Close
 
 'Date
 itmModif.SubItems(I_COL_MODIF_DATE) = rstProjSoum.Fields("Date")
 
 'Heure
 itmModif.SubItems(I_COL_MODIF_HEURE) = rstProjSoum.Fields("Heure")
 
 Call rstProjSoum.MoveNext
 Loop
40 End If
 
Call rstProjSoum.Close
 
4 Call rstProjSoum.Open("SELECT * FROM " & sTable & " WHERE " & sChamps & " = '" & txtNoProjSoum.Text & "' AND Type = 'E' AND TypeModif = 'FACTURATION' ORDER BY [Date], Heure", g_connData, adOpenDynamic, adLockOptimistic)
 
4 If Not rstProjSoum.EOF Then
 'Ajout de la section "Modifications"
4 Set itmModif = lvwHistorique.ListItems.Add
 
4 itmModif.Text = "Facturation"
 
4 itmModif.Bold = True
 
4 Do While Not rstProjSoum.EOF
 'Ajout des modifications
4 Set itmModif = lvwHistorique.ListItems.Add
 
4 Call rstEmploye.Open("SELECT Employe FROM GrbEmployés WHERE noEmploye = " & rstProjSoum.Fields("NoEmployé"), g_connData, adOpenDynamic, adLockOptimistic)
 
 'Employé
4 itmModif.Text = rstEmploye.Fields("Employe")
 
4 Call rstEmploye.Close
 
 'Date
4  itmModif.SubItems(I_COL_MODIF_DATE) = rstProjSoum.Fields("Date")
 
 'Heure
4  itmModif.SubItems(I_COL_MODIF_HEURE) = rstProjSoum.Fields("Heure")

 'Montant
4  itmModif.SubItems(I_COL_MODIF_MONTANT) = rstProjSoum.Fields("Valeur")
 
4  Call rstProjSoum.MoveNext
4  Loop
4  End If
 
4  Call rstProjSoum.Close
4  Set rstProjSoum = Nothing

50 Set rstEmploye = Nothing

50 Exit Sub

Oups:

 wOups "frmProjSoumElec", "RemplirListViewModifications", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListViewSuppression()

 On Error GoTo Oups

 'Rempli le listView avec les pièces supprimées
 Dim rstBavard As ADODB.Recordset
 Dim rstEmploye As ADODB.Recordset
 Dim itmBavard As ListItem

 Call lvwBavard.ListItems.Clear

 Set rstBavard = New ADODB.Recordset
 Set rstEmploye = New ADODB.Recordset

 Call rstBavard.Open("SELECT * FROM GrbBavardSuppression WHERE NoProjSoum = '" & txtNoProjSoum.Text & "' AND Type = 'E' ORDER BY [Date], Heure", g_connData, adOpenDynamic, adLockOptimistic)

 Do While Not rstBavard.EOF
 Set itmBavard = lvwBavard.ListItems.Add

 Call rstEmploye.Open("SELECT Employe FROM GrbEmployés WHERE NoEmploye = " & rstBavard.Fields("IDUser"), g_connData, adOpenDynamic, adLockOptimistic)

  itmBavard.Text = rstEmploye.Fields("Employe")

  Call rstEmploye.Close

  itmBavard.SubItems(I_COL_SUPP_DATE) = rstBavard.Fields("Date")
  itmBavard.SubItems(I_COL_SUPP_HEURE) = rstBavard.Fields("Heure")
  itmBavard.SubItems(I_COL_SUPP_QTE) = rstBavard.Fields("Qté")
  itmBavard.SubItems(I_COL_SUPP_NO_ITEM) = rstBavard.Fields("No Item")

  Call rstBavard.MoveNext
  Loop

10 Call rstBavard.Close
Set rstBavard = Nothing

Set rstEmploye = Nothing

Exit Sub

Oups:

wOups "frmProjSoumElec", "RemplirListViewSuppression", Err, Err.number, Err.Description
End Sub

Private Sub Cmdajouter_Click()

 On Error GoTo Oups

 'Ajoute une soumission ou un projet
 Dim rstProjSoum As ADODB.Recordset
 Dim sNumero As String
 Dim sNoProjet As String
 Dim bExiste As Boolean
 Dim bProjet As Boolean
 Dim bContinuer As Boolean
 Dim bNoValide As Boolean
 
 'Affiche le message de saisie selon le type
 If m_eType = TYPE_PROJET Then
 sNumero = InputBox("Veuillez entrer le numéro du projet")
 Else
  If MsgBox("Voulez-vous créer une nouvelle soumission?" & vbNewLine & _
 "Oui - Nouvelle soumission" & vbNewLine & _
 "Non - Copie d'un projet dans une soumission", vbYesNo) = vbYes Then
  sNumero = InputBox("Veuillez entrer le numéro de la soumission")
  Else
  sNumero = InputBox("Veuillez entrer le numéro de la soumission")

  sNoProjet = InputBox("À partir de quel projet voulez-vous créer cette soumission?")

  bProjet = True
  End If
  End If
 
10 If bProjet = True Then
1 If sNumero <> vbNullString And sNoProjet <> vbNullString Then
 bContinuer = True
 End If
Else
 If sNumero <> vbNullString Then
 bContinuer = True
 End If
End If
 
If bContinuer = True Then
 Screen.MousePointer = vbHourglass

 bNoValide = True

If ValiderFormatNumeroProjSoum(sNumero) = False Then
 bNoValide = False
 End If

 If bNoValide = True Then
 If ValiderFormatElectrique(sNumero) = False Then
 bNoValide = False
 End If
1  End If

 If bNoValide = True Then
 If m_eType = TYPE_PROJET Then
 If ValiderFormatJobSansSoum(sNumero) = False Then
 bNoValide = False
 End If
 Else
 If ValiderFormatSoumission(sNumero) = False Then
 bNoValide = False
 End If
 End If
 End If

 If bNoValide = False Then
 Screen.MousePointer = vbDefault

 Exit Sub
End If

 sNumero = UCase(sNumero)

Set rstProjSoum = New ADODB.Recordset

 'Ouvre le recordset selon le type
 Call rstProjSoum.Open("SELECT * FROM GrbSoumissionElec WHERE IDSoumission = '" & sNumero & "'", g_connData, adOpenDynamic, adLockPessimistic)
 
If rstProjSoum.EOF Then
 bExiste = False
Else
bExiste = True

 Call MsgBox("Le numéro " & sNumero & " existe dans les soumissions électriques!", vbOKOnly, "Erreur")
 End If

 Call rstProjSoum.Close

 If bExiste = False Then
 Call rstProjSoum.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = '" & sNumero & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 If rstProjSoum.EOF Then
 bExiste = False
 Else
 bExiste = True

 Call MsgBox("Le numéro " & sNumero & " existe dans les projets électriques!", vbOKOnly, "Erreur")
 End If

 Call rstProjSoum.Close
End If

 If bExiste = False Then
 Call rstProjSoum.Open("SELECT * FROM GrbSoumissionMec WHERE IDSoumission = '" & sNumero & "'", g_connData, adOpenDynamic, adLockOptimistic)

 If rstProjSoum.EOF Then
 bExiste = False
 Else
 bExiste = True

4 Call MsgBox("Le numéro " & sNumero & " existe dans les soumissions mécaniques!", vbOKOnly, "Erreur")
4 End If

4 Call rstProjSoum.Close
4 End If
 
4 If bExiste = False Then
4 Call rstProjSoum.Open("SELECT * FROM GrbProjetMec WHERE IDProjet = '" & sNumero & "'", g_connData, adOpenDynamic, adLockOptimistic)

4 If rstProjSoum.EOF Then
4 bExiste = False
4 Else
4 bExiste = True

4 Call MsgBox("Le numéro " & sNumero & " existe dans les projets mécaniques!", vbOKOnly, "Erreur")
4  End If

4  Call rstProjSoum.Close
4  End If

 'Si le projet ou la soumission n'existe pas
4  If bExiste = False Then
 'Pour pouvoir réafficher l'élément affiché avant l'ajout au cas où la personne
 'annule l'ajout
4  Call rstProjSoum.Open("SELECT * FROM GrbProjSoum WHERE IDProjSoum = '" & sNumero & "'", g_connData, adOpenDynamic, adLockOptimistic)

4  If Not rstProjSoum.EOF Then
4  If rstProjSoum.Fields("Ouvert") = False Then
4  Call MsgBox("Ce numéro est fermé!", vbOKOnly, "Erreur")

50 Call rstProjSoum.Close
 Set rstProjSoum = Nothing

 Screen.MousePointer = vbDefault

 Exit Sub
 End If
 End If

 Call rstProjSoum.Close
 Set rstProjSoum = Nothing
 
 If bProjet = False Then
 Call InitialiserTempsTaux(True)
 
 Call InitialiserNouveauxTaux
 
 m_sAncienProjSoum = txtNoProjSoum.Text
 
 'Affiche le nouveau numéro
5  txtNoProjSoum.Text = sNumero
 
5  Call InitialiserVariables(txtNoProjSoum.Text)
 
 'Débarre les champs
5  Call BarrerChamps(False)
 
 'Vide les champs
5  Call ViderChamps
5  Else
5  If VerifierProjet(sNoProjet) = True Then
 'Débarre les champs
5  Call BarrerChamps(False)
 
 'Vide les champs
5  Call ViderChamps

60 txtNoProjSoum.Text = sNumero

  Call RemplirSoumissionProjet(sNumero, sNoProjet)
  Else
  Call MsgBox("Le projet " & sNoProjet & " n'existe pas!", vbOKOnly, "Erreur")

  Screen.MousePointer = vbDefault

  Exit Sub
  End If
  End If

 'Vide la valeur par défaut si demande Sous-Section
  m_sSousSection = vbNullString
 
  m_bModeAjout = True
  m_bModeAffichage = False
 
  lvwSoumission.Height = lvwSoumission.Height * 0.49
6  lvwSoumission.Top = lvwPieces.Top + lvwPieces.Height + (lvwSoumission.Height * 0.02)
 
 'Met le form en mode ajout/modif
6  Call AfficherControles(MODE_AJOUT_MODIF)
6  End If
6  End If
 
6  Screen.MousePointer = vbDefault

6  Exit Sub

Oups:

6  wOups "frmProjSoumElec", "cmdAjouter_Click", Err, Err.number, Err.Description
End Sub

Private Function VerifierProjet(ByVal sNoProjet As String) As Boolean
 
 On Error GoTo Oups

 Dim rstProjet As ADODB.Recordset
 Dim bExiste As Boolean

 Set rstProjet = New ADODB.Recordset

 Call rstProjet.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)

 If Not rstProjet.EOF Then
 bExiste = True
 Else
 bExiste = False
 End If

 Call rstProjet.Close
  Set rstProjet = Nothing

  VerifierProjet = bExiste

  Exit Function
 
Oups:

  wOups "frmProjSoumElec", "VerifierProjet", Err, Err.number, Err.Description
End Function

Private Sub RemplirSoumissionProjet(ByVal sNoSoumission As String, ByVal sNoProjet As String)

 On Error GoTo Oups
 
 'Affiche le projet ou la soumission choisie
 Dim rstProjSoum As ADODB.Recordset
 Dim rstConfig As ADODB.Recordset
 Dim bVariables As Boolean
 Dim bTauxHoraire As Boolean
 Dim bPrixPieces As Boolean

 Set rstProjSoum = New ADODB.Recordset
 Set rstConfig = New ADODB.Recordset

 If MsgBox("Voulez-vous mettre à jour les variables systèmes?" & vbNewLine & _
 "- % Profit" & vbNewLine & _
 "- % Commission" & vbNewLine & _
 "- % Imprévu" & vbNewLine & _
 "- $ Pages manuel", vbYesNo) = vbYes Then
 bVariables = True
 Else
  bVariables = False
  End If

  If MsgBox("Voulez-vous mettre à jour les taux horaires?", vbYesNo) = vbYes Then
  bTauxHoraire = True
  Else
  bTauxHoraire = False
  End If

  If MsgBox("Voulez-vous mettre à jour le prix des pièces?", vbYesNo) = vbYes Then
bPrixPieces = True
Else
 bPrixPieces = False
End If
 
Call rstProjSoum.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = '" & sNoProjet & "'", g_connData, adOpenDynamic, adLockOptimistic)

m_bSansTemps = rstProjSoum.Fields("SansTemps")

If Not IsNull(rstProjSoum.Fields("TempsDessin")) Then
 m_sTempsDessin = rstProjSoum.Fields("TempsDessin")
Else
 m_sTempsDessin = "0"
End If

If Not IsNull(rstProjSoum.Fields("TempsFabrication")) Then
m_sTempsFabrication = rstProjSoum.Fields("TempsFabrication")
Else
 m_sTempsFabrication = "0"
End If

 If Not IsNull(rstProjSoum.Fields("TempsAssemblage")) Then
 m_sTempsAssemblage = rstProjSoum.Fields("TempsAssemblage")
 Else
1  m_sTempsAssemblage = "0"
 End If

 If Not IsNull(rstProjSoum.Fields("TempsProgInterface")) Then
 m_sTempsProgInterface = rstProjSoum.Fields("TempsProgInterface")
Else
 m_sTempsProgInterface = "0"
End If

If Not IsNull(rstProjSoum.Fields("TempsProgAutomate")) Then
 m_sTempsProgAutomate = rstProjSoum.Fields("TempsProgAutomate")
Else
 m_sTempsProgAutomate = "0"
End If

If Not IsNull(rstProjSoum.Fields("TempsProgRobot")) Then
m_sTempsProgRobot = rstProjSoum.Fields("TempsProgRobot")
Else
m_sTempsProgRobot = "0"
End If

2  If Not IsNull(rstProjSoum.Fields("TempsVision")) Then
 m_sTempsVision = rstProjSoum.Fields("TempsVision")
2  Else
 m_sTempsVision = "0"
30 End If

If Not IsNull(rstProjSoum.Fields("TempsTest")) Then
 m_sTempsTest = rstProjSoum.Fields("TempsTest")
Else
 m_sTempsTest = "0"
End If

If Not IsNull(rstProjSoum.Fields("TempsInstallation")) Then
 m_sTempsInstallation = rstProjSoum.Fields("TempsInstallation")
Else
 m_sTempsInstallation = "0"
End If

If Not IsNull(rstProjSoum.Fields("TempsMiseService")) Then
m_sTempsMiseService = rstProjSoum.Fields("TempsMiseService")
Else
m_sTempsMiseService = "0"
End If

3  If Not IsNull(rstProjSoum.Fields("TempsFormation")) Then
 m_sTempsFormation = rstProjSoum.Fields("TempsFormation")
3  Else
 m_sTempsFormation = "0"
40 End If

If Not IsNull(rstProjSoum.Fields("TempsGestion")) Then
4 m_sTempsGestion = rstProjSoum.Fields("TempsGestion")
4 Else
4 m_sTempsGestion = "0"
4 End If

4 If Not IsNull(rstProjSoum.Fields("TempsShipping")) Then
4 m_sTempsShipping = rstProjSoum.Fields("TempsShipping")
4 Else
4 m_sTempsShipping = "0"
4 End If
 
4 Call rstConfig.Open("SELECT * FROM GrbConfig", g_connData, adOpenForwardOnly, adLockReadOnly)

4  If bTauxHoraire = True Then
4  If Not IsNull(rstConfig.Fields("TauxDessinElec")) Then
4  m_sTauxDessin = rstConfig.Fields("TauxDessinElec")
4  Else
4  m_sTauxDessin = "0"
4  End If

4  If Not IsNull(rstConfig.Fields("TauxFabrication")) Then
4  m_sTauxFabrication = rstConfig.Fields("TauxFabrication")
50 Else
m_sTauxFabrication = "0"
 End If

 If Not IsNull(rstConfig.Fields("TauxAssemblageElec")) Then
 m_sTauxAssemblage = rstConfig.Fields("TauxAssemblageElec")
 Else
 m_sTauxAssemblage = "0"
 End If

 If Not IsNull(rstConfig.Fields("TauxProgInterface")) Then
 m_sTauxProgInterface = rstConfig.Fields("TauxProgInterface")
 Else
 m_sTauxProgInterface = "0"
5  End If

5  If Not IsNull(rstConfig.Fields("TauxProgAutomate")) Then
5  m_sTauxProgAutomate = rstConfig.Fields("TauxProgAutomate")
5  Else
5  m_sTauxProgAutomate = "0"
5  End If

5  If Not IsNull(rstConfig.Fields("TauxProgRobot")) Then
5  m_sTauxProgRobot = rstConfig.Fields("TauxProgRobot")
60 Else
  m_sTauxProgRobot = "0"
  End If

  If Not IsNull(rstConfig.Fields("TauxVision")) Then
  m_sTauxVision = rstConfig.Fields("TauxVision")
  Else
  m_sTauxVision = "0"
  End If

  If Not IsNull(rstConfig.Fields("TauxTestElec")) Then
  m_sTauxTest = rstConfig.Fields("TauxTestElec")
  Else
  m_sTauxTest = "0"
6  End If

6  If Not IsNull(rstConfig.Fields("TauxInstallationElec")) Then
6  m_sTauxInstallation = rstConfig.Fields("TauxInstallationElec")
6  Else
6  m_sTauxInstallation = "0"
6  End If

6  If Not IsNull(rstConfig.Fields("TauxMiseService")) Then
6  m_sTauxMiseService = rstConfig.Fields("TauxMiseService")
70 Else
  m_sTauxMiseService = "0"
  End If

  If Not IsNull(rstConfig.Fields("TauxFormationElec")) Then
  m_sTauxFormation = rstConfig.Fields("TauxFormationElec")
  Else
  m_sTauxFormation = "0"
  End If

  If Not IsNull(rstConfig.Fields("TauxGestionProjetsElec")) Then
  m_sTauxGestion = rstConfig.Fields("TauxGestionProjetsElec")
  Else
  m_sTauxGestion = "0"
   End If

   If Not IsNull(rstConfig.Fields("TauxShippingElec")) Then
7  m_sTauxShipping = rstConfig.Fields("TauxShippingElec")
7  Else
7  m_sTauxShipping = "0"
7  End If
7  Else
7  If Not IsNull(rstProjSoum.Fields("TauxDessin")) Then
80 m_sTauxDessin = rstProjSoum.Fields("TauxDessin")
  Else
  m_sTauxDessin = "0"
  End If

  If Not IsNull(rstProjSoum.Fields("TauxFabrication")) Then
  m_sTauxFabrication = rstProjSoum.Fields("TauxFabrication")
  Else
  m_sTauxFabrication = "0"
  End If

  If Not IsNull(rstProjSoum.Fields("TauxAssemblage")) Then
  m_sTauxAssemblage = rstProjSoum.Fields("TauxAssemblage")
  Else
   m_sTauxAssemblage = "0"
   End If

   If Not IsNull(rstProjSoum.Fields("TauxProgInterface")) Then
   m_sTauxProgInterface = rstProjSoum.Fields("TauxProgInterface")
8  Else
8  m_sTauxProgInterface = "0"
8  End If

8  If Not IsNull(rstProjSoum.Fields("TauxProgAutomate")) Then
90 m_sTauxProgAutomate = rstProjSoum.Fields("TauxProgAutomate")
  Else
  m_sTauxProgAutomate = "0"
  End If

  If Not IsNull(rstProjSoum.Fields("TauxProgRobot")) Then
  m_sTauxProgRobot = rstProjSoum.Fields("TauxProgRobot")
  Else
  m_sTauxProgRobot = "0"
  End If

  If Not IsNull(rstProjSoum.Fields("TauxVision")) Then
  m_sTauxVision = rstProjSoum.Fields("TauxVision")
  Else
 m_sTauxVision = "0"
   End If

 If Not IsNull(rstProjSoum.Fields("TauxTest")) Then
   m_sTauxTest = rstProjSoum.Fields("TauxTest")
 Else
   m_sTauxTest = "0"
 End If

9  If Not IsNull(rstProjSoum.Fields("TauxInstallation")) Then
 m_sTauxInstallation = rstProjSoum.Fields("TauxInstallation")
10Else
1 m_sTauxInstallation = "0"
1End If

 If Not IsNull(rstProjSoum.Fields("TauxMiseService")) Then
1 m_sTauxMiseService = rstProjSoum.Fields("TauxMiseService")
 Else
1 m_sTauxMiseService = "0"
 End If

1 If Not IsNull(rstProjSoum.Fields("TauxFormation")) Then
 m_sTauxFormation = rstProjSoum.Fields("TauxFormation")
1Else
10  m_sTauxFormation = "0"
10  End If

10  If Not IsNull(rstProjSoum.Fields("TauxGestion")) Then
10  m_sTauxGestion = rstProjSoum.Fields("TauxGestion")
10  Else
10  m_sTauxGestion = "0"
10  End If

10  If Not IsNull(rstProjSoum.Fields("TauxShipping")) Then
1 m_sTauxShipping = rstProjSoum.Fields("TauxShipping")
11Else
1 m_sTauxShipping = "0"
1 End If
11 End If

11 If bVariables = True Then
1 m_sProfit = rstConfig.Fields("ProfitElec")
1 m_sCommission = rstConfig.Fields("Commission")
1 m_sImprevue = rstConfig.Fields("Imprévus")
11 Else
1 m_sProfit = rstProjSoum.Fields("Profit")
1 m_sCommission = rstProjSoum.Fields("Commission")
11  m_sImprevue = rstProjSoum.Fields("Imprevue")
11  End If

1 Call rstConfig.Close
11  Set rstConfig = Nothing
 
1 txtProjet.Text = rstProjSoum.Fields("Description")
11  txtNbreManuel.Text = rstProjSoum.Fields("NbreManuel")
1 txtPrixManuel.Text = rstProjSoum.Fields("total_manuel")
11  txtTransport.Text = rstProjSoum.Fields("transport")

1 If Not IsNull(rstProjSoum.Fields("CheminPhotos")) Then
1 txtCheminPhotos.Text = rstProjSoum.Fields("CheminPhotos")
12 Else
1 txtCheminPhotos.Text = vbNullString
12 End If
 
12 chkCSA.Value = Abs(CInt(rstProjSoum.Fields("CSA")))
12 chkCUL.Value = Abs(CInt(rstProjSoum.Fields("CUL")))
12 chkUL.Value = Abs(CInt(rstProjSoum.Fields("UL")))
12 chkCUR.Value = Abs(CInt(rstProjSoum.Fields("CUR")))
12 chkUR.Value = Abs(CInt(rstProjSoum.Fields("UR")))
12 chkCE.Value = Abs(CInt(rstProjSoum.Fields("CE")))

12 txtPrixTotal.Text = rstProjSoum.Fields("PrixTotal")
12  txtProfit.Text = rstProjSoum.Fields("total_profit")

12  If Not IsNull(rstProjSoum.Fields("Delais")) Then
12  txtDelais.Text = rstProjSoum.Fields("Delais")
12  Else
12  txtDelais.Text = "0"
12  End If

12  txtCommission.Text = rstProjSoum.Fields("total_commission")

12  If Not IsNull(rstProjSoum.Fields("MontantForfait")) Then
130 txtForfait.Text = rstProjSoum.Fields("MontantForfait")

13 If Not IsNull(rstProjSoum.Fields("InitialeForfait")) Then
1 lblForfaitInitiale.Caption = rstProjSoum.Fields("InitialeForfait")
1 Else
1 lblForfaitInitiale.Caption = ""
1 End If
13 Else
1 txtForfait.Text = ""
1 lblForfaitInitiale.Caption = ""
13 End If

13 Call rstProjSoum.Close
13 Set rstProjSoum = Nothing
 
 'Affiche les pieces de la soumission
13  Call RemplirListViewSoumissionProjet(sNoProjet)

13  If bPrixPieces = True Then
13  Call UpdatePieces
13  End If

13  m_bModeAffichage = False

1 Call CalculerPrix

139Exit Sub

Oups:

1 wOups "frmProjSoumElec", "RemplirProjSoum", Err, Err.number, Err.Description
End Sub

Private Sub RechercherProjSoum(ByVal sNoProjSoum As String)

 On Error GoTo Oups

 'Méthode qui recherche une soumission ou un projet dans le combo
 'et qui le sélectionne
 Dim iCompteur As Integer
 
 'Pour chaque élément du combo
 For iCompteur = 0 To cmbProjSoum.ListCount - 1
 'Si le texte de l'élément du combo est égal au numéro recherché
 If cmbProjSoum.LIST(iCompteur) = sNoProjSoum Then
 'On le sélectionne
 cmbProjSoum.ListIndex = iCompteur
 
 Exit For
 End If
 Next

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "RechercherProjSoum", Err, Err.number, Err.Description
End Sub

Private Sub RemplirProjSoum()

 On Error GoTo Oups

 'Affiche le projet ou la soumission choisie
 Dim rstProjSoum As ADODB.Recordset
 Dim rstSoum As ADODB.Recordset
 Dim rstClient As ADODB.Recordset
 Dim rstContact As ADODB.Recordset

 Set rstProjSoum = New ADODB.Recordset
 Set rstSoum = New ADODB.Recordset
 Set rstClient = New ADODB.Recordset
 Set rstContact = New ADODB.Recordset
 
 'Ouvre le recordset selon le type
 If m_eType = TYPE_PROJET Then
 Call rstProjSoum.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
  If Not IsNull(rstProjSoum.Fields("IDSoumission")) Then
  txtNoSoumission.Text = rstProjSoum.Fields("IDSoumission")
  Else
  txtNoSoumission.Text = vbNullString
  End If

  If Right$(txtNoProjSoum.Text, 2) >= "60" And Right$(txtNoProjSoum.Text, 2) <= "98" Then
  If Trim(rstProjSoum.Fields("LiaisonChargeable")) <> "" Then
  m_sLiaison = rstProjSoum.Fields("LiaisonChargeable")
 Else
 m_sLiaison = vbNullString

 Do While Trim$(m_sLiaison) = vbNullString
 m_sLiaison = InputBox("Quelle est l'extention au projet " & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 3) & " auquel ce projet sera lié?")
 Loop

 rstProjSoum.Fields("LiaisonChargeable") = m_sLiaison

 Call rstProjSoum.Update
 End If
 End If
Else
 Call rstProjSoum.Open("SELECT * FROM GrbSoumissionElec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
End If

1  m_bSansTemps = rstProjSoum.Fields("SansTemps")
 
 'Recordset pour avoir le nom du client
Call rstClient.Open("SELECT NomClient, IDClient FROM GrbClient WHERE IDClient = " & rstProjSoum.Fields("IDClient"), g_connData, adOpenDynamic, adLockOptimistic)
 
 'Recordset pour avoir le nom du contact
 Call rstContact.Open("SELECT NomContact, IDContact FROM GrbContact WHERE IDContact = " & rstProjSoum.Fields("IDContact"), g_connData, adOpenDynamic, adLockOptimistic)
 
txtClient.Text = rstClient.Fields("NomClient")
 
 txtcontact.Text = rstContact.Fields("NomContact")
 
Call rstClient.Close
 Set rstClient = Nothing
 
1  Call rstContact.Close
 Set rstContact = Nothing
 
 txtProjet.Text = rstProjSoum.Fields("Description")
txtNbreManuel.Text = rstProjSoum.Fields("NbreManuel")
txtPrixManuel.Text = Conversion(rstProjSoum.Fields("total_manuel"), MODE_PAS_FORMAT)
txtTransport.Text = rstProjSoum.Fields("transport")
 
txtTotalPieces.Text = Conversion(rstProjSoum.Fields("Total_Piece"), MODE_ARGENT)
txtTotalTemps.Text = Conversion(rstProjSoum.Fields("Total_Temps"), MODE_ARGENT)
txtPrixTotal.Text = Conversion(rstProjSoum.Fields("PrixTotal"), MODE_ARGENT)
txtImprevus.Text = Conversion(rstProjSoum.Fields("Total_Imprevue"), MODE_ARGENT)
txtProfit.Text = Conversion(rstProjSoum.Fields("total_profit"), MODE_ARGENT)
txtCommission.Text = Conversion(rstProjSoum.Fields("total_commission"), MODE_ARGENT)

If Not IsNull(rstProjSoum.Fields("CheminPhotos")) Then
txtCheminPhotos.Text = rstProjSoum.Fields("CheminPhotos")
Else
txtCheminPhotos.Text = vbNullString
End If
 
2  chkCSA.Value = Abs(CInt(rstProjSoum.Fields("CSA")))
chkCUL.Value = Abs(CInt(rstProjSoum.Fields("CUL")))
2  chkUL.Value = Abs(CInt(rstProjSoum.Fields("UL")))
chkCUR.Value = Abs(CInt(rstProjSoum.Fields("CUR")))
30 chkUR.Value = Abs(CInt(rstProjSoum.Fields("UR")))
chkCE.Value = Abs(CInt(rstProjSoum.Fields("CE")))
 
If Not IsNull(rstProjSoum.Fields("Delais")) Then
 txtDelais.Text = Trim(rstProjSoum.Fields("Delais"))
Else
 txtDelais.Text = ""
End If

If Not IsNull(rstProjSoum.Fields("MontantForfait")) Then
 txtForfait.Text = Conversion(rstProjSoum.Fields("MontantForfait"), MODE_ARGENT)

 If Not IsNull(rstProjSoum.Fields("InitialeForfait")) Then
 lblForfaitInitiale.Caption = rstProjSoum.Fields("InitialeForfait")
 Else
 lblForfaitInitiale.Caption = ""
 End If
3  Else
 txtForfait.Text = ""
lblForfaitInitiale.Caption = ""
End If

3  If m_eType = TYPE_PROJET Then
 If Not IsNull(rstProjSoum.Fields("PrixRéception")) Then
 If Trim(rstProjSoum.Fields("PrixRéception")) <> "" Then
4 txtPrixReception.Text = Conversion(rstProjSoum.Fields("PrixRéception"), MODE_ARGENT)
4 Else
4 txtPrixReception.Text = Conversion("0", MODE_ARGENT)
4 End If
4 Else
4 txtPrixReception.Text = Conversion("0", MODE_ARGENT)
4 End If

4 If Not IsNull(rstProjSoum.Fields("IDSoumission")) Then
4 Call rstSoum.Open("SELECT * FROM GrbSoumissionElec WHERE IDSoumission = '" & rstProjSoum.Fields("IDSoumission") & "'", g_connData, adOpenDynamic, adLockOptimistic)

4 If Not rstSoum.EOF Then
4 If Not IsNull(rstSoum.Fields("PrixTotal")) Then
4  If Trim(rstSoum.Fields("PrixTotal")) <> "" Then
4  txtPrixSoumission.Text = Conversion(rstSoum.Fields("PrixTotal"), MODE_ARGENT)
4  Else
4  txtPrixSoumission.Text = Conversion(0, MODE_ARGENT)
4  End If
4  Else
4  txtPrixSoumission.Text = Conversion(0, MODE_ARGENT)
4  End If
50 Else
 txtPrixSoumission.Text = Conversion(0, MODE_ARGENT)
 End If

 Call rstSoum.Close
 Set rstSoum = Nothing
 End If
 End If

 Call rstProjSoum.Close
 Set rstProjSoum = Nothing
 
 'Affiche les pieces de la soumission
 Call RemplirListViewProjSoum(txtNoProjSoum.Text)

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "RemplirProjSoum", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboCategoriesPieces()

 On Error GoTo Oups

 'Remplir le combo des tables (Pièces)
 Dim rstCatalogueElec As ADODB.Recordset
 
 'Il faut vider le combo avant de le remplir
 Call cmbPieces.Clear
 
 'Cette méthode crée un recordset contenant les categorie
 'le nom de toutes les tables de la BD
 Set rstCatalogueElec = New ADODB.Recordset
 
 Call rstCatalogueElec.Open("SELECT DISTINCT CATEGORIE FROM GrbCatalogueElec ORDER BY CATEGORIE", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Tant que ce n'est pas la fin des enregistrements
 Do While Not rstCatalogueElec.EOF
 Call cmbPieces.AddItem(rstCatalogueElec.Fields("CATEGORIE"))
 
 Call rstCatalogueElec.MoveNext
 Loop
 
 Call rstCatalogueElec.Close
 Set rstCatalogueElec = Nothing
 
  If cmbPieces.ListCount > 0 Then
  cmbPieces.ListIndex = 0
  End If

  Exit Sub

Oups:

  wOups "frmProjSoumElec", "RemplirComboCategoriesPieces", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboClients(ByVal sRecherche As String)

 On Error GoTo Oups

 'Remplit le combo des clients
 Dim rstClient As ADODB.Recordset
 
 Call cmbclient.Clear
 
 Set rstClient = New ADODB.Recordset
 
 Call rstClient.Open("SELECT NomClient, IDClient FROM GrbClient WHERE Instr(1, NomClient, '" & Replace(sRecherche, "'", "''") & "') > 0 AND Supprimé = False", g_connData, adOpenDynamic, adLockOptimistic)
 
 Do While Not rstClient.EOF
 'on met le nom du client dans le combo
 Call cmbclient.AddItem(rstClient.Fields("NomClient"))
 
 'on met l'id du client dans l'itemdata
 cmbclient.ItemData(cmbclient.newIndex) = rstClient.Fields("IDClient")
 
 Call rstClient.MoveNext
 Loop
 
 Call rstClient.Close
  Set rstClient = Nothing

  Exit Sub

Oups:

  wOups "frmProjSoumElec", "RemplirComboClients", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboContacts()

 On Error GoTo Oups

 'Remplis le combo des contacts selon le client choisi
 Dim rstContact As ADODB.Recordset
 
 Call cmbContact.Clear
 
 If cmbclient.ListIndex > -1 Then
 Set rstContact = New ADODB.Recordset

 Call rstContact.Open("SELECT NomContact, IDContact FROM GrbContact INNER JOIN GrbContactClient ON GrbContact.IDContact = GrbContactClient.NoContact WHERE GrbContactClient.noClient = " & cmbclient.ItemData(cmbclient.ListIndex), g_connData, adOpenDynamic, adLockOptimistic)
 
 'Si il n'y a aucun contact pour le client choisi
 If rstContact.EOF Then
 'On ajoute tous les contacts
 Call rstContact.Close
 
 Call rstContact.Open("SELECT NomContact, IDContact FROM GrbContact WHERE Supprimé = False", g_connData, adOpenDynamic, adLockOptimistic)
 End If
 
 Do While Not rstContact.EOF
 'On ajoute le nom du contact dans le combo
  Call cmbContact.AddItem(rstContact.Fields("NomContact"))
 
 'On ajoute l'id du contact dans l'itemdata du combo
  cmbContact.ItemData(cmbContact.newIndex) = rstContact.Fields("IDContact")
 
  Call rstContact.MoveNext
  Loop
 
  Call rstContact.Close
  Set rstContact = Nothing
 
 'Si le combo n'est pas vide, on sélectionne le premier élément
  If cmbContact.ListCount > 0 Then
  cmbContact.ListIndex = 0
End If
End If

Exit Sub

Oups:

wOups "frmProjSoumElec", "RemplirComboContacts", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboSections()

 On Error GoTo Oups
 
 'Remplis le combo des sections
 Dim rstSection As ADODB.Recordset
 Dim sChamps As String
 
 Call cmbSections.Clear
 
 Set rstSection = New ADODB.Recordset
 
 'Il faut le remplir selon l'ordre
 Call rstSection.Open("SELECT * FROM GrbSoumProjSectionElec ORDER BY Ordre", g_connData, adOpenDynamic, adLockOptimistic)
 
 If m_eLangage = ANGLAIS Then
 sChamps = "NomSectionEN"
 Else
 sChamps = "NomSectionFR"
 End If
 
  Do While Not rstSection.EOF
 'On met le nom de la section dans le combo
  If Not IsNull(rstSection.Fields(sChamps)) Then
  Call cmbSections.AddItem(rstSection.Fields(sChamps))
  Else
  Call cmbSections.AddItem(vbNullString)
  End If
 
 'On met l'id de la section dans l'itemdata du combo
  cmbSections.ItemData(cmbSections.newIndex) = rstSection.Fields("IDSection")
 
  Call rstSection.MoveNext
10 Loop
 
Call rstSection.Close
Set rstSection = Nothing
 
 'Si le combo n'est pas vide, on sélectionne le premier élément
If cmbSections.ListCount > 0 Then
 cmbSections.ListIndex = 0
End If

Exit Sub

Oups:

wOups "frmProjSoumElec", "RemplirComboSections", Err, Err.number, Err.Description
End Sub

Private Sub cmdImprimer_Click()

 On Error GoTo Oups

 Dim rstProjSoum As ADODB.Recordset
 Dim sUser As String

 If txtNoProjSoum.Text <> vbNullString Then
 If VerifierSiOuvert(sUser) = False Then

 Set rstProjSoum = New ADODB.Recordset
 
 'Ouvre les tables
 If m_eType = TYPE_PROJET Then
 If MsgBox("Voulez-vous faire imprimer le projet et tous les extras associés à ce projet?", vbYesNo) = vbYes Then
 Call rstProjSoum.Open("SELECT * FROM GrbProjetElec WHERE Left(IDProjet, 6) = '" & Left$(txtNoProjSoum.Text, 6) & "' ORDER BY IDProjet", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstProjSoum.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = '" & txtNoProjSoum.Text & "' ORDER BY IDProjet", g_connData, adOpenDynamic, adLockOptimistic)
  End If
  Else
  Call rstProjSoum.Open("SELECT * FROM GrbSoumissionElec WHERE IDSoumission = '" & txtNoProjSoum.Text & "' ORDER BY IDSoumission", g_connData, adOpenDynamic, adLockOptimistic)
  End If

   bTrigger = False
7  intdummie = 0

 '***********************************************************************************
 'AJOUT PAR GAÉTAN GINGRAS 0  FÉVRIER 2010
 '***********************************************************************************
 If MsgBox("Désirez-vous afficher les dates de réception et de commande?", vbYesNo, "Date de réception et de commande") = vbYes Then
 bFlag = True
 Else
 bFlag = False
 End If
 '***********************************************************************************
 
  Do While Not rstProjSoum.EOF
  If m_eType = TYPE_PROJET Then
  Call CalculerTotalRecordset(rstProjSoum.Fields("IDProjet"))
  End If

 Call ImprimerProjSoum(rstProjSoum)
 If Not intdummie = vbYes Then
 Call ImprimerListePieces(rstProjSoum)
10  End If

 Call rstProjSoum.MoveNext
 Loop

 Call rstProjSoum.Close
 Set rstProjSoum = Nothing
 Else
 If m_eType = TYPE_PROJET Then
 Call MsgBox("Ce projet est en modification par " & sUser & "!", vbOKOnly, "Erreur")
 Else
 Call MsgBox("Cette soumission est en modification par " & sUser & "!", vbOKOnly, "Erreur")
 End If
End If
End If

 Exit Sub

Oups:

wOups "frmProjSoumElec", "cmdImprimer_Click", Err, Err.number, Err.Description
End Sub

Private Sub ImprimerProjSoum(ByVal rstProjSoum As ADODB.Recordset)

 On Error GoTo Oups

 'Impression de la feuille de soumission
 Dim rstPiece As ADODB.Recordset
 Dim rstPrixSoum As ADODB.Recordset
 Dim rstTemp As ADODB.Recordset
 Dim rstImpProjSoum As ADODB.Recordset
 Dim rstSoum As ADODB.Recordset
 Dim sOrdreSection As String
 Dim iCompteurSoum As Integer
 Dim sSousSection As String
 Dim sSousSectionRS As String
 Dim dblTempsDessin As Double
  Dim dblTempsFabrication As Double
  Dim dblTempsAssemblage As Double
  Dim dblTempsProgInterface As Double
  Dim dblTempsProgAutomate As Double
  Dim dblTempsProgRobot As Double
  Dim dblTempsVision As Double
  Dim dblTempsTest As Double
  Dim dblTempsInstallation As Double
10 Dim dblTempsMiseService As Double
Dim dblTempsFormation As Double
Dim dblTempsGestion As Double
Dim dblTempsShipping As Double
Dim dblTotalTemps As Double
Dim dblTotalAutre As Double
Dim dblTotalReste As Double
Dim dblTotalHebergement As Double
Dim dblTotalRepas As Double
Dim dblTotalTransport As Double
Dim dblTotalUniteMobile As Double
Dim sChampsSection As String
1  Dim sNoProjet As String
Dim sNoSoumission As String
 Dim dblPrixEmballage As Double
 
 'Supprime les données de l'impression
Call g_connData.Execute("DELETE * FROM Grbimpression_soumission")
 
 iCompteurSoum = 1
 
Screen.MousePointer = vbHourglass

 Set rstImpProjSoum = New ADODB.Recordset

1  Call rstImpProjSoum.Open("SELECT * FROM Grbimpression_soumission", g_connData, adOpenDynamic, adLockOptimistic)
 
 sOrdreSection = vbNullString
 
 Set rstPiece = New ADODB.Recordset
 
If m_eType = TYPE_PROJET Then
 sNoProjet = rstProjSoum.Fields("IDProjet")

 If Not IsNull(rstProjSoum.Fields("IDSoumission")) Then
 sNoSoumission = rstProjSoum.Fields("IDSoumission")
 Else
 sNoSoumission = vbNullString
 End If

 Call rstPiece.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & sNoProjet & "' And Type = 'E' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
Else
 sNoProjet = vbNullString
sNoSoumission = rstProjSoum.Fields("IDSoumission")

 Call rstPiece.Open("SELECT * FROM GrbSoumission_Pieces WHERE IDSoumission = '" & sNoSoumission & "' And Type = 'E' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
2  End If
 
Set rstTemp = New ADODB.Recordset
 
2  Do While Not rstPiece.EOF
 sSousSectionRS = rstPiece.Fields("SousSection")
 
If sSousSectionRS = S_PAS_SOUS_SECTION Then
 sSousSectionRS = " "
End If
 
3 If sOrdreSection <> rstPiece.Fields("OrdreSection") Then
 'remplis la table impression_soumission
 'ajoute seulement la section
 sOrdreSection = rstPiece.Fields("OrdreSection")
 
 If m_eLangage = ANGLAIS Then
 sChampsSection = "NomSectionEN"
 Else
 sChampsSection = "NomSectionFR"
 End If

 Call rstTemp.Open("SELECT " & sChampsSection & " FROM GrbSoumProjSectionElec WHERE IDSection = " & rstPiece.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
 
 'Ajoute la section dans la soumission
 Call rstImpProjSoum.AddNew
 
 rstImpProjSoum.Fields("NoLigne") = iCompteurSoum

 If m_eType = TYPE_PROJET Then
 rstImpProjSoum.Fields("IDSoumission") = sNoProjet
 Else
 rstImpProjSoum.Fields("IDSoumission") = sNoSoumission
 End If
 
 If Not IsNull(rstTemp.Fields(sChampsSection)) Then
 rstImpProjSoum.Fields("NomSection") = rstTemp.Fields(sChampsSection)
 Else
 rstImpProjSoum.Fields("NomSection") = " "
 End If
 
4 Call rstImpProjSoum.Update
 
4 iCompteurSoum = iCompteurSoum + 1
 
4 Call rstTemp.Close
 
4 sSousSection = rstPiece.Fields("SousSection")
 
4 If sSousSection = S_PAS_SOUS_SECTION Or sSousSection = "" Then
4 sSousSection = " "
4 End If
 
4 Call rstImpProjSoum.AddNew

4 rstImpProjSoum.Fields("NoLigne") = iCompteurSoum

4 If m_eType = TYPE_PROJET Then
4 rstImpProjSoum.Fields("IDSoumission") = sNoProjet
4  Else
4  rstImpProjSoum.Fields("IDSoumission") = sNoSoumission
4  End If

4  rstImpProjSoum.Fields("SousSection") = sSousSection
 
4  Call rstImpProjSoum.Update
 
4  iCompteurSoum = iCompteurSoum + 1
4  Else
 'ajoute une soussection dans impression_soum
4  If sSousSection <> sSousSectionRS Then
50 sSousSection = sSousSectionRS
 
 Call rstImpProjSoum.AddNew
 
 rstImpProjSoum.Fields("NoLigne") = iCompteurSoum

 If m_eType = TYPE_PROJET Then
 rstImpProjSoum.Fields("IDSoumission") = sNoProjet
 Else
 rstImpProjSoum.Fields("IDSoumission") = sNoSoumission
 End If

 rstImpProjSoum.Fields("SousSection") = sSousSectionRS
 
 Call rstImpProjSoum.Update
 
 iCompteurSoum = iCompteurSoum + 1
 End If
5  End If
 
 'ajoute une piece dans impression_soum
5  Call rstImpProjSoum.AddNew
 
5  rstImpProjSoum.Fields("NoLigne") = iCompteurSoum

5  If m_eType = TYPE_PROJET Then
5  rstImpProjSoum.Fields("IDSoumission") = sNoProjet
5  Else
5  rstImpProjSoum.Fields("IDSoumission") = sNoSoumission
5  End If

60 rstImpProjSoum.Fields("NumItem") = rstPiece.Fields("NumItem")
  rstImpProjSoum.Fields("Qté") = rstPiece.Fields("Qté")
 
  If m_eLangage = ANGLAIS Then
  rstImpProjSoum.Fields("Description") = rstPiece.Fields("DESC_EN")
  Else
  rstImpProjSoum.Fields("Description") = rstPiece.Fields("DESC_FR")
  End If


 '************************************************************************************************
 'SECTION MODIFIER PAR GAÉTAN GINGRAS LE   FÉVRIER 2010
 '************************************************************************************************
 
  rstImpProjSoum.Fields("MANUFACT") = rstPiece.Fields("MANUFACT")
  'rstImpProjSoum.Fields("PRIX_LIST") = rstPiece.Fields("PRIX_LIST")
 
  'If Trim(rstPiece.Fields("ESCOMPTE")) <> vbNullString Then
  ' rstImpProjSoum.Fields("ESCOMPTE") = Conversion(Replace(rstPiece.Fields("ESCOMPTE"), ".", ","), MODE_POURCENT)
  'Else
6  ' rstImpProjSoum.Fields("ESCOMPTE") = Conversion(0, MODE_POURCENT)
6  'End If
 
6  rstImpProjSoum.Fields("PRIX_NET") = rstPiece.Fields("PRIX_NET")
 
6  Call rstTemp.Open("SELECT NomFournisseur FROM GrbFournisseur WHERE IDFRS = " & rstPiece.Fields("IDFRS"), g_connData, adOpenDynamic, adLockOptimistic)
 
6  If Not rstTemp.EOF Then
6  rstImpProjSoum.Fields("NomFournisseur") = rstTemp.Fields("NomFournisseur")
6  End If
 
6  Call rstTemp.Close

70 'rstImpProjSoum.Fields("TEMPS") = rstPiece.Fields("TEMPS")
  'rstImpProjSoum.Fields("TEMPS_TOTAL") = rstPiece.Fields("TEMPS_TOTAL")
  rstImpProjSoum.Fields("PRIX_TOTAL") = rstPiece.Fields("PRIX_TOTAL")
71 rstImpProjSoum.Fields("PROFIT_ARGENT") = rstPiece.Fields("PROFIT_ARGENT")


 ''''''''''''''''''''''''''''''''''''''''''''''''''''
 '' ajout du Numéros sequentiel de commande dans impression
 ''''''''''''''''''''''''''''''''
 rstImpProjSoum.Fields("NoSéquentiel") = rstPiece.Fields("NoSéquentiel")

 'AJOUT DE CETTE SECTION PAR GAÉTAN GINGRAS LE   FÉVRIER 2010
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
 Else 'il n'y a pas de champ date de réception et de commande dans la table GrbSoumission_Pièces
 rstImpProjSoum.Fields("DateReception") = ""
 rstImpProjSoum.Fields("DateCommande") = ""
 End If
 '************************************************************************************************
 'FIN DE LA SECTION DE MODIFICATION
 '************************************************************************************************
 
71 Call rstImpProjSoum.Update
 
714 iCompteurSoum = iCompteurSoum + 1
 
 'prochaine enreg
  Call rstPiece.MoveNext
  Loop
 
 'ferme les tables
  Call rstImpProjSoum.Close
 
 ''''''''''''''''''''''''''''''''''
 ' rapport soumission, met dans l'ordre de ligne
 ''''''''''''''''''''''''''''''''''''
 
  Dim sProjet As String

71  If m_eType = TYPE_PROJET Then
  sProjet = sNoProjet
72 Else
72 sProjet = sNoSoumission
72 End If

724 Call rstImpProjSoum.Open("SELECT * FROM Grbimpression_Soumission WHERE IDSoumission = '" & sProjet & "' ORDER BY NoLigne", g_connData, adOpenDynamic, adLockOptimistic)

  Set DR_SoumissionElec.DataSource = rstImpProjSoum

 '***********************************************************************
 'Pour cas d'urgence, la fonction d'exporter dans Excel va être ici.
 'Nous pourrons plus tard le mettre à un meilleur endroit.
 'Gaétan Gingras le 14 mai 2009
 '***********************************************************************

  If bTrigger = False Then
  bTrigger = True
  intdummie = MsgBox("Désirez-vous exporter les données dans Excel, SEULEMENT ?", vbYesNo + vbInformation, "Exportation dans Excel")
  End If
  If intdummie = vbYes Then
74 Dim sqlstr As String
74 Dim rstExport As ADODB.Recordset
744 Set rstExport = New ADODB.Recordset
  'sqlstr = "SELECT Grbimpression_soumission.IDSoumission, CDbl([Qté]) AS Quantité, Grbimpression_soumission.NumItem, Grbimpression_soumission.Description, Grbimpression_soumission.Manufact, CDbl([Prix_list]) AS PrixdeListe, CDbl(Left([escompte],Len([escompte])-1)) AS Escomptes, CDbl([Prix_net]) AS prix_nette, Grbimpression_soumission.NomFournisseur, Grbimpression_soumission.DateReception , Grbimpression_soumission.DateCommande "
74  sqlstr = "SELECT Grbimpression_soumission.IDSoumission, CDbl([Qté]) AS Quantité, Grbimpression_soumission.NumItem, Grbimpression_soumission.Description, Grbimpression_soumission.Manufact, CDbl([Prix_list]) AS PrixdeListe, CDbl(Left([escompte],Len([escompte])-1)) AS Escomptes, CDbl([Prix_net]) AS prix_nette, Grbimpression_soumission.Prix_total - Grbimpression_soumission.Profit_Argent AS Prix_Total ,Grbimpression_soumission.NomFournisseur, Grbimpression_soumission.DateReception , Grbimpression_soumission.DateCommande , Grbimpression_soumission.NoSéquentiel "

 sqlstr = sqlstr + "FROM Grbimpression_soumission "
74  sqlstr = sqlstr + "WHERE (((Grbimpression_soumission.IDSoumission)='" & sProjet & "') AND ((Grbimpression_soumission.NumItem) Is Not Null)) "
74  sqlstr = sqlstr + "ORDER BY Grbimpression_soumission.noligne"
74  Call rstExport.Open(sqlstr, g_connData, adOpenDynamic, adLockOptimistic)
  Call ExportdansExcel(rstExport)
75 Screen.MousePointer = vbDefault
75 Exit Sub
75 End If
 '***********************************************************************
 
7  Call TraduireImpressionSoumission

7  If m_eType = TYPE_PROJET Then
7  DR_SoumissionElec.Sections("Section5").Controls("shpCadrePrixReception").Visible = True
7  DR_SoumissionElec.Sections("Section5").Controls("lblTitrePrixReception").Visible = True
7  DR_SoumissionElec.Sections("Section5").Controls("lblPrixReception").Visible = True

80 DR_SoumissionElec.Sections("Section5").Controls("shpCadrePrixSoumission").Visible = True
  DR_SoumissionElec.Sections("Section5").Controls("lblTitrePrixSoumission").Visible = True
  DR_SoumissionElec.Sections("Section5").Controls("lblPrixSoumission").Visible = True
  Else
  DR_SoumissionElec.Sections("Section5").Controls("shpCadrePrixReception").Visible = False
  DR_SoumissionElec.Sections("Section5").Controls("lblTitrePrixReception").Visible = False
  DR_SoumissionElec.Sections("Section5").Controls("lblPrixReception").Visible = False

  DR_SoumissionElec.Sections("Section5").Controls("shpCadrePrixSoumission").Visible = False
  DR_SoumissionElec.Sections("Section5").Controls("lblTitrePrixSoumission").Visible = False
  DR_SoumissionElec.Sections("Section5").Controls("lblPrixSoumission").Visible = False
  End If
 
 'affiche la date
 '**************************************************
 'ajout par Gaétan Gingras le 20 mai 2009
854 If MsgBox("Désirez-vous afficher la date en bas de page ?", vbYesNo + vbInformation, "Affichage de la date") = vbYes Then
  DR_SoumissionElec.Sections("section3").Controls("lbldate").Caption = ConvertDate(Date)
85  Else
85  DR_SoumissionElec.Sections("section3").Controls("lbldate").Caption = " "
85  End If
 '**************************************************
 
 'affiche entete
   If Not IsNull(rstProjSoum.Fields("IDSoumission")) Then
   DR_SoumissionElec.Sections("Section2").Controls("lblSoumission").Caption = rstProjSoum.Fields("IDSoumission")
   Else
   DR_SoumissionElec.Sections("Section2").Controls("lblSoumission").Caption = vbNullString
8  End If
 
8  If m_eType = TYPE_PROJET Then
8  DR_SoumissionElec.Sections("Section2").Controls("lblprojet").Caption = rstProjSoum.Fields("IDProjet")
8  Else
90 DR_SoumissionElec.Sections("Section2").Controls("lblProjet").Caption = vbNullString
90 End If
 
  DR_SoumissionElec.Sections("Section2").Controls("lbldescription").Caption = rstProjSoum.Fields("Description")
 
  Call rstTemp.Open("SELECT NomClient FROM GrbClient WHERE IDClient = " & rstProjSoum.Fields("IDClient"), g_connData, adOpenDynamic, adLockOptimistic)
 
  DR_SoumissionElec.Sections("Section2").Controls("lblclient").Caption = rstTemp.Fields("NomClient")

  Call rstTemp.Close
 
  Call rstTemp.Open("SELECT NomContact FROM GrbContact WHERE IDContact = " & rstProjSoum.Fields("IDContact"), g_connData, adOpenDynamic, adLockOptimistic)
 
  DR_SoumissionElec.Sections("Section2").Controls("lblcontact").Caption = rstTemp.Fields("NomContact")
 
  Call rstTemp.Close
 
 'Affiche pied d'état
 
 'Temps
  If Not IsNull(rstProjSoum.Fields("TauxDessin")) Then
  DR_SoumissionElec.Sections("Section5").Controls("lblTauxDessin").Caption = rstProjSoum.Fields("TauxDessin")
  Else
 DR_SoumissionElec.Sections("Section5").Controls("lblTauxDessin").Caption = "0"
   End If

 If Not IsNull(rstProjSoum.Fields("TauxFabrication")) Then
   DR_SoumissionElec.Sections("Section5").Controls("lblTauxFabrication").Caption = rstProjSoum.Fields("TauxFabrication")
 Else
   DR_SoumissionElec.Sections("Section5").Controls("lblTauxFabrication").Caption = "0"
 End If

9  If Not IsNull(rstProjSoum.Fields("TauxAssemblage")) Then
 DR_SoumissionElec.Sections("Section5").Controls("lblTauxAssemblage").Caption = rstProjSoum.Fields("TauxAssemblage")
100 Else
1DR_SoumissionElec.Sections("Section5").Controls("lblTauxAssemblage").Caption = "0"
10 End If

If Not IsNull(rstProjSoum.Fields("TauxProgInterface")) Then
1DR_SoumissionElec.Sections("Section5").Controls("lblTauxProgInterface").Caption = rstProjSoum.Fields("TauxProgInterface")
Else
1DR_SoumissionElec.Sections("Section5").Controls("lblTauxProgInterface").Caption = "0"
End If

10 If Not IsNull(rstProjSoum.Fields("TauxProgAutomate")) Then
 DR_SoumissionElec.Sections("Section5").Controls("lblTauxProgAutomate").Caption = rstProjSoum.Fields("TauxProgAutomate")
10 Else
10  DR_SoumissionElec.Sections("Section5").Controls("lblTauxProgAutomate").Caption = "0"
10  End If

10  If Not IsNull(rstProjSoum.Fields("TauxProgRobot")) Then
10  DR_SoumissionElec.Sections("Section5").Controls("lblTauxProgRobot").Caption = rstProjSoum.Fields("TauxProgRobot")
108Else
10  DR_SoumissionElec.Sections("Section5").Controls("lblTauxProgRobot").Caption = "0"
109End If

10  If Not IsNull(rstProjSoum.Fields("TauxVision")) Then
110 DR_SoumissionElec.Sections("Section5").Controls("lblTauxVision").Caption = rstProjSoum.Fields("TauxVision")
110 Else
1 DR_SoumissionElec.Sections("Section5").Controls("lblTauxVision").Caption = "0"
11 End If

11 If Not IsNull(rstProjSoum.Fields("TauxTest")) Then
1 DR_SoumissionElec.Sections("Section5").Controls("lblTauxTest").Caption = rstProjSoum.Fields("TauxTest")
11 Else
1 DR_SoumissionElec.Sections("Section5").Controls("lblTauxTest").Caption = "0"
11 End If

11 If Not IsNull(rstProjSoum.Fields("TauxInstallation")) Then
1 DR_SoumissionElec.Sections("Section5").Controls("lblTauxInstallation").Caption = rstProjSoum.Fields("TauxInstallation")
11 Else
11  DR_SoumissionElec.Sections("Section5").Controls("lblTauxInstallation").Caption = "0"
11  End If

1 If Not IsNull(rstProjSoum.Fields("TauxMiseService")) Then
1 DR_SoumissionElec.Sections("Section5").Controls("lblTauxMiseService").Caption = rstProjSoum.Fields("TauxMiseService")
1 Else
1 DR_SoumissionElec.Sections("Section5").Controls("lblTauxMiseService").Caption = "0"
1 End If

11  If Not IsNull(rstProjSoum.Fields("TauxFormation")) Then
 DR_SoumissionElec.Sections("Section5").Controls("lblTauxFormation").Caption = rstProjSoum.Fields("TauxFormation")
1Else
1 DR_SoumissionElec.Sections("Section5").Controls("lblTauxFormation").Caption = "0"
12 End If

12 If Not IsNull(rstProjSoum.Fields("TauxGestion")) Then
1 DR_SoumissionElec.Sections("Section5").Controls("lblTauxGestion").Caption = rstProjSoum.Fields("TauxGestion")
12 Else
1 DR_SoumissionElec.Sections("Section5").Controls("lblTauxGestion").Caption = "0"
12 End If

12 If Not IsNull(rstProjSoum.Fields("TauxShipping")) Then
1 DR_SoumissionElec.Sections("Section5").Controls("lblTauxShipping").Caption = rstProjSoum.Fields("TauxShipping")
12 Else
12  DR_SoumissionElec.Sections("Section5").Controls("lblTauxShipping").Caption = "0"
12  End If

12  If Not IsNull(rstProjSoum.Fields("TempsDessin")) Then
1 DR_SoumissionElec.Sections("Section5").Controls("lblTempsDessinSoum").Caption = rstProjSoum.Fields("TempsDessin")
128Else
1 DR_SoumissionElec.Sections("Section5").Controls("lblTempsDessinSoum").Caption = "0"
129End If

12  If Not IsNull(rstProjSoum.Fields("TempsFabrication")) Then
130 If rstProjSoum.Fields("SansTemps") = False Then
13 DR_SoumissionElec.Sections("Section5").Controls("lblTempsFabricationSoum").Caption = rstProjSoum.Fields("TempsFabrication")
1 Else
1 DR_SoumissionElec.Sections("Section5").Controls("lblTempsFabricationSoum").Caption = "0"
1 End If
13 Else
1 DR_SoumissionElec.Sections("Section5").Controls("lblTempsFabricationSoum").Caption = "0"
13 End If

13 If Not IsNull(rstProjSoum.Fields("TempsAssemblage")) Then
1 DR_SoumissionElec.Sections("Section5").Controls("lblTempsAssemblageSoum").Caption = rstProjSoum.Fields("TempsAssemblage")
13 Else
1 DR_SoumissionElec.Sections("Section5").Controls("lblTempsAssemblageSoum").Caption = "0"
136End If

13  If Not IsNull(rstProjSoum.Fields("TempsProgInterface")) Then
13  DR_SoumissionElec.Sections("Section5").Controls("lblTempsProgInterfaceSoum").Caption = rstProjSoum.Fields("TempsProgInterface")
13  Else
1DR_SoumissionElec.Sections("Section5").Controls("lblTempsProgInterfaceSoum").Caption = "0"
1End If

13  If Not IsNull(rstProjSoum.Fields("TempsProgAutomate")) Then
1 DR_SoumissionElec.Sections("Section5").Controls("lblTempsProgAutomateSoum").Caption = rstProjSoum.Fields("TempsProgAutomate")
140Else
14DR_SoumissionElec.Sections("Section5").Controls("lblTempsProgAutomateSoum").Caption = "0"
14 End If

14 If Not IsNull(rstProjSoum.Fields("TempsProgRobot")) Then
14 DR_SoumissionElec.Sections("Section5").Controls("lblTempsProgRobotSoum").Caption = rstProjSoum.Fields("TempsProgRobot")
14 Else
14 DR_SoumissionElec.Sections("Section5").Controls("lblTempsProgRobotSoum").Caption = "0"
14 End If

14 If Not IsNull(rstProjSoum.Fields("TempsVision")) Then
14 DR_SoumissionElec.Sections("Section5").Controls("lblTempsVisionSoum").Caption = rstProjSoum.Fields("TempsVision")
14 Else
14 DR_SoumissionElec.Sections("Section5").Controls("lblTempsVisionSoum").Caption = "0"
146End If

14  If Not IsNull(rstProjSoum.Fields("TempsTest")) Then
14  DR_SoumissionElec.Sections("Section5").Controls("lblTempsTestSoum").Caption = rstProjSoum.Fields("TempsTest")
14  Else
14  DR_SoumissionElec.Sections("Section5").Controls("lblTempsTestSoum").Caption = "0"
14  End If

14  If Not IsNull(rstProjSoum.Fields("TempsInstallation")) Then
14  DR_SoumissionElec.Sections("Section5").Controls("lblTempsInstallationSoum").Caption = rstProjSoum.Fields("TempsInstallation")
150Else
15DR_SoumissionElec.Sections("Section5").Controls("lblTempsInstallationSoum").Caption = "0"
End If

1 If Not IsNull(rstProjSoum.Fields("TempsMiseService")) Then
 DR_SoumissionElec.Sections("Section5").Controls("lblTempsMiseServiceSoum").Caption = rstProjSoum.Fields("TempsMiseService")
Else
 DR_SoumissionElec.Sections("Section5").Controls("lblTempsMiseServiceSoum").Caption = "0"
End If

If Not IsNull(rstProjSoum.Fields("TempsFormation")) Then
 DR_SoumissionElec.Sections("Section5").Controls("lblTempsFormationSoum").Caption = rstProjSoum.Fields("TempsFormation")
Else
 DR_SoumissionElec.Sections("Section5").Controls("lblTempsFormationSoum").Caption = "0"
156End If

15  If Not IsNull(rstProjSoum.Fields("TempsGestion")) Then
15  DR_SoumissionElec.Sections("Section5").Controls("lblTempsGestionSoum").Caption = rstProjSoum.Fields("TempsGestion")
15  Else
15  DR_SoumissionElec.Sections("Section5").Controls("lblTempsGestionSoum").Caption = "0"
15  End If

15  If Not IsNull(rstProjSoum.Fields("TempsShipping")) Then
15  DR_SoumissionElec.Sections("Section5").Controls("lblTempsShippingSoum").Caption = rstProjSoum.Fields("TempsShipping")
160Else
16DR_SoumissionElec.Sections("Section5").Controls("lblTempsShippingSoum").Caption = "0"
1  End If

1  If Not IsNull(rstProjSoum.Fields("TempsDessin")) Then
 If IsNumeric(rstProjSoum.Fields("TempsDessin")) Then
 dblTempsDessin = CDbl(rstProjSoum.Fields("TempsDessin"))
 Else
 dblTempsDessin = 0
 End If
1  Else
 dblTempsDessin = 0
1  End If

16  If Not IsNull(rstProjSoum.Fields("TempsFabrication")) Then
16  If rstProjSoum.Fields("SansTemps") = False Then
16  If IsNumeric(rstProjSoum.Fields("TempsFabrication")) Then
16  dblTempsFabrication = CDbl(rstProjSoum.Fields("TempsFabrication"))
16  Else
16  dblTempsFabrication = 0
16  End If
16  Else
170 dblTempsFabrication = 0
 End If
1  Else
 dblTempsFabrication = 0
1  End If

1  If Not IsNull(rstProjSoum.Fields("TempsAssemblage")) Then
 If IsNumeric(rstProjSoum.Fields("TempsAssemblage")) Then
 dblTempsAssemblage = CDbl(rstProjSoum.Fields("TempsAssemblage"))
 Else
 dblTempsAssemblage = 0
 End If
1  Else
1   dblTempsAssemblage = 0
1   End If

17  If Not IsNull(rstProjSoum.Fields("TempsProgInterface")) Then
17  If IsNumeric(rstProjSoum.Fields("TempsProgInterface")) Then
17  dblTempsProgInterface = CDbl(rstProjSoum.Fields("TempsProgInterface"))
17  Else
17  dblTempsProgInterface = 0
17  End If
180Else
 dblTempsProgInterface = 0
1  End If

1  If Not IsNull(rstProjSoum.Fields("TempsProgAutomate")) Then
 If IsNumeric(rstProjSoum.Fields("TempsProgAutomate")) Then
 dblTempsProgAutomate = CDbl(rstProjSoum.Fields("TempsProgAutomate"))
 Else
 dblTempsProgAutomate = 0
 End If
1  Else
 dblTempsProgAutomate = 0
1  End If

1   If Not IsNull(rstProjSoum.Fields("TempsProgRobot")) Then
1   If IsNumeric(rstProjSoum.Fields("TempsProgRobot")) Then
1   dblTempsProgRobot = CDbl(rstProjSoum.Fields("TempsProgRobot"))
1   Else
18  dblTempsProgRobot = 0
18  End If
189Else
18  dblTempsProgRobot = 0
190End If

If Not IsNull(rstProjSoum.Fields("TempsVision")) Then
1  If IsNumeric(rstProjSoum.Fields("TempsVision")) Then
1  dblTempsVision = CDbl(rstProjSoum.Fields("TempsVision"))
1  Else
1  dblTempsVision = 0
1  End If
1  Else
1  dblTempsVision = 0
1  End If

1  If Not IsNull(rstProjSoum.Fields("TempsTest")) Then
1  If IsNumeric(rstProjSoum.Fields("TempsTest")) Then
 dblTempsTest = CDbl(rstProjSoum.Fields("TempsTest"))
1   Else
 dblTempsTest = 0
1   End If
 Else
1   dblTempsTest = 0
 End If

19  If Not IsNull(rstProjSoum.Fields("TempsInstallation")) Then
200 If IsNumeric(rstProjSoum.Fields("TempsInstallation")) Then
 dblTempsInstallation = CDbl(rstProjSoum.Fields("TempsInstallation"))
 Else
 dblTempsInstallation = 0
 End If
Else
 dblTempsInstallation = 0
End If

If Not IsNull(rstProjSoum.Fields("TempsMiseService")) Then
 If IsNumeric(rstProjSoum.Fields("TempsMiseService")) Then
 dblTempsMiseService = CDbl(rstProjSoum.Fields("TempsMiseService"))
 Else
20  dblTempsMiseService = 0
20  End If
207Else
20  dblTempsMiseService = 0
208End If

20  If Not IsNull(rstProjSoum.Fields("TempsFormation")) Then
20  If IsNumeric(rstProjSoum.Fields("TempsFormation")) Then
20  dblTempsFormation = CDbl(rstProjSoum.Fields("TempsFormation"))
210 Else
21 dblTempsFormation = 0
2 End If
21 Else
2 dblTempsFormation = 0
21 End If

21 If Not IsNull(rstProjSoum.Fields("TempsGestion")) Then
2 If IsNumeric(rstProjSoum.Fields("TempsGestion")) Then
dblTempsGestion = CDbl(rstProjSoum.Fields("TempsGestion"))
2 Else
dblTempsGestion = 0
2 End If
216Else
2 dblTempsGestion = 0
2 End If

21  If Not IsNull(rstProjSoum.Fields("TempsShipping")) Then
 If IsNumeric(rstProjSoum.Fields("TempsShipping")) Then
dblTempsShipping = CDbl(rstProjSoum.Fields("TempsShipping"))
 Else
21  dblTempsShipping = 0
 End If
2Else
2 dblTempsShipping = 0
22 End If

22 dblTotalTemps = dblTempsDessin + _
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
 
22 DR_SoumissionElec.Sections("Section5").Controls("lblTotalTempsRHSoum").Caption = dblTotalTemps

22 If m_eType = TYPE_PROJET Then
2 Call CalculerTempsReelsImpression(rstProjSoum.Fields("IDProjet"))
22 End If
 
 'Autres frais
22 If m_eType = TYPE_PROJET Then
2 DR_SoumissionElec.Sections("Section5").Controls("lblNbrePersonne").Caption = "0"
2 DR_SoumissionElec.Sections("Section5").Controls("lblTempsHebergement").Caption = "0"
22  DR_SoumissionElec.Sections("Section5").Controls("lblTauxHebergement1").Caption = "0"
2 DR_SoumissionElec.Sections("Section5").Controls("lblTauxHebergement2").Caption = "0"
22  DR_SoumissionElec.Sections("Section5").Controls("lblTempsRepas").Caption = "0"
2 DR_SoumissionElec.Sections("Section5").Controls("lblTauxRepas").Caption = "0"
22  DR_SoumissionElec.Sections("Section5").Controls("lblTempsTransport").Caption = "0"
2 DR_SoumissionElec.Sections("Section5").Controls("lblTauxTransport").Caption = "0"
22  DR_SoumissionElec.Sections("Section5").Controls("lblTempsUniteMobile").Caption = "0"
2 DR_SoumissionElec.Sections("Section5").Controls("lblTauxUniteMobile").Caption = "0"
230Else
23 If Not IsNull(rstProjSoum.Fields("NbrePersonne")) Then
2 DR_SoumissionElec.Sections("Section5").Controls("lblNbrePersonne").Caption = rstProjSoum.Fields("NbrePersonne")
2 Else
2 DR_SoumissionElec.Sections("Section5").Controls("lblNbrePersonne").Caption = "0"
2 End If

2 If Not IsNull(rstProjSoum.Fields("TempsHebergement")) Then
2 DR_SoumissionElec.Sections("Section5").Controls("lblTempsHebergement").Caption = rstProjSoum.Fields("TempsHebergement")
2 Else
2 DR_SoumissionElec.Sections("Section5").Controls("lblTempsHebergement").Caption = "0"
2 End If

2 If Not IsNull(rstProjSoum.Fields("TauxHebergement1")) Then
2 DR_SoumissionElec.Sections("Section5").Controls("lblTauxHebergement1").Caption = rstProjSoum.Fields("TauxHebergement1")
2 Else
2 DR_SoumissionElec.Sections("Section5").Controls("lblTauxHebergement1").Caption = "0"
2 End If

2 If Not IsNull(rstProjSoum.Fields("TauxHebergement2")) Then
2 DR_SoumissionElec.Sections("Section5").Controls("lblTauxHebergement2").Caption = rstProjSoum.Fields("TauxHebergement2")
2Else
2 DR_SoumissionElec.Sections("Section5").Controls("lblTauxHebergement2").Caption = "0"
240 End If
 
24 If Not IsNull(rstProjSoum.Fields("TempsRepas")) Then
24 DR_SoumissionElec.Sections("Section5").Controls("lblTempsRepas").Caption = rstProjSoum.Fields("TempsRepas")
24 Else
24 DR_SoumissionElec.Sections("Section5").Controls("lblTempsRepas").Caption = "0"
24 End If

24 If Not IsNull(rstProjSoum.Fields("TauxRepas")) Then
24 DR_SoumissionElec.Sections("Section5").Controls("lblTauxRepas").Caption = rstProjSoum.Fields("TauxRepas")
24 Else
24 DR_SoumissionElec.Sections("Section5").Controls("lblTauxRepas").Caption = "0"
24 End If

24 If Not IsNull(rstProjSoum.Fields("TempsTransport")) Then
24  DR_SoumissionElec.Sections("Section5").Controls("lblTempsTransport").Caption = rstProjSoum.Fields("TempsTransport")
24  Else
24  DR_SoumissionElec.Sections("Section5").Controls("lblTempsTransport").Caption = "0"
24  End If

24  If Not IsNull(rstProjSoum.Fields("TauxTransport")) Then
24  DR_SoumissionElec.Sections("Section5").Controls("lblTauxTransport").Caption = rstProjSoum.Fields("TauxTransport")
24  Else
24  DR_SoumissionElec.Sections("Section5").Controls("lblTauxTransport").Caption = "0"
250 End If

25 If Not IsNull(rstProjSoum.Fields("TempsUniteMobile")) Then
 DR_SoumissionElec.Sections("Section5").Controls("lblTempsUniteMobile").Caption = rstProjSoum.Fields("TempsUniteMobile")
2 Else
 DR_SoumissionElec.Sections("Section5").Controls("lblTempsUniteMobile").Caption = "0"
2 End If

 If Not IsNull(rstProjSoum.Fields("TauxUniteMobile")) Then
 DR_SoumissionElec.Sections("Section5").Controls("lblTauxUniteMobile").Caption = rstProjSoum.Fields("TauxUniteMobile")
 Else
 DR_SoumissionElec.Sections("Section5").Controls("lblTauxUniteMobile").Caption = "0"
 End If
End If

25  If Not IsNull(rstProjSoum.Fields("PrixEmballage")) Then
25  DR_SoumissionElec.Sections("Section5").Controls("lblPrixEmballage").Caption = rstProjSoum.Fields("PrixEmballage")
257Else
25  DR_SoumissionElec.Sections("Section5").Controls("lblPrixEmballage").Caption = "0"
258End If

25  DR_SoumissionElec.Sections("Section5").Controls("lblPrixManuel").Caption = rstProjSoum.Fields("Total_manuel")

259DR_SoumissionElec.Sections("Section5").Controls("lblTotalTempsAR").Caption = Conversion(rstProjSoum.Fields("total_temps"), MODE_ARGENT)
25  DR_SoumissionElec.Sections("Section5").Controls("lblTotalPieceAR").Caption = Conversion(rstProjSoum.Fields("total_piece"), MODE_ARGENT)
260DR_SoumissionElec.Sections("Section5").Controls("lblProfit").Caption = Conversion((rstProjSoum.Fields("profit") - 1) * 100, MODE_POURCENT)
260 DR_SoumissionElec.Sections("Section5").Controls("lblTotalProfit").Caption = Conversion(rstProjSoum.Fields("total_profit"), MODE_ARGENT)
2  DR_SoumissionElec.Sections("Section5").Controls("lblImprevue").Caption = Conversion(rstProjSoum.Fields("imprevue"), MODE_POURCENT)
2  DR_SoumissionElec.Sections("Section5").Controls("lblImprevueAR").Caption = Conversion(rstProjSoum.Fields("total_imprevue"), MODE_ARGENT)
2  DR_SoumissionElec.Sections("Section5").Controls("lblCommission").Caption = Conversion(rstProjSoum.Fields("commission"), MODE_POURCENT)
2  DR_SoumissionElec.Sections("Section5").Controls("lblCommissionAR").Caption = Conversion(rstProjSoum.Fields("total_commission"), MODE_ARGENT)

2  If m_eType = TYPE_PROJET Then
 If Not IsNull(rstProjSoum.Fields("PrixRéception")) Then
 DR_SoumissionElec.Sections("Section5").Controls("lblPrixReception").Caption = Conversion(rstProjSoum.Fields("PrixRéception"), MODE_ARGENT)
 Else
 DR_SoumissionElec.Sections("Section5").Controls("lblPrixReception").Caption = Conversion(0, MODE_ARGENT)
 End If

26  If Not IsNull(rstProjSoum.Fields("IDSoumission")) Then
26  Set rstPrixSoum = New ADODB.Recordset

26  Call rstPrixSoum.Open("SELECT PrixTotal FROM GrbSoumissionElec WHERE IDSoumission = '" & rstProjSoum.Fields("IDSoumission") & "'", g_connData, adOpenDynamic, adLockOptimistic)

26  If Not rstPrixSoum.EOF Then
26  If Not IsNull(rstPrixSoum.Fields("PrixTotal")) Then
26  DR_SoumissionElec.Sections("Section5").Controls("lblPrixSoumission").Caption = Conversion(rstPrixSoum.Fields("PrixTotal"), MODE_ARGENT)
26  Else
26  DR_SoumissionElec.Sections("Section5").Controls("lblPrixSoumission").Caption = Conversion("0", MODE_ARGENT)
270 End If
2  Else
 DR_SoumissionElec.Sections("Section5").Controls("lblPrixSoumission").Caption = Conversion("0", MODE_ARGENT)
 End If

 Call rstPrixSoum.Close
 Set rstPrixSoum = Nothing
 Else
 DR_SoumissionElec.Sections("Section5").Controls("lblPrixSoumission").Caption = Conversion("0", MODE_ARGENT)
 End If
2  End If
 
2  If m_eType = TYPE_PROJET Then
 dblTotalHebergement = 0
2   dblTotalRepas = 0
2   dblTotalTransport = 0
27  dblTotalUniteMobile = 0
27  Else
27  If Not IsNull(rstProjSoum.Fields("TotalHebergement")) Then
27  dblTotalHebergement = rstProjSoum.Fields("TotalHebergement")
27  End If
 
27  If Not IsNull(rstProjSoum.Fields("TotalRepas")) Then
280 dblTotalRepas = rstProjSoum.Fields("TotalRepas")
28End If
 
 If Not IsNull(rstProjSoum.Fields("TempsTransport")) And Not IsNull(rstProjSoum.Fields("TauxTransport")) Then
 dblTotalTransport = CDbl(rstProjSoum.Fields("TempsTransport")) * CDbl(rstProjSoum.Fields("TauxTransport"))
 Else
 dblTotalTransport = 0
 End If

 If Not IsNull(rstProjSoum.Fields("TempsUniteMobile")) And Not IsNull(rstProjSoum.Fields("TauxUniteMobile")) Then
 dblTotalUniteMobile = CDbl(rstProjSoum.Fields("TempsUniteMobile")) * CDbl(rstProjSoum.Fields("TauxUniteMobile"))
 Else
 dblTotalUniteMobile = 0
 End If
286End If

2   If Not IsNull(rstProjSoum.Fields("PrixEmballage")) Then
2   dblPrixEmballage = CDbl(Replace(rstProjSoum.Fields("PrixEmballage"), ".", ","))
2   Else
28  dblPrixEmballage = 0
28  End If
 
289dblTotalReste = dblTotalHebergement + dblTotalRepas + dblTotalTransport + dblTotalUniteMobile + dblPrixEmballage

28  dblTotalAutre = dblTotalReste + CDbl(rstProjSoum.Fields("total_manuel"))
 
290DR_SoumissionElec.Sections("Section5").Controls("lblAutre").Caption = Conversion(CStr(dblTotalAutre), MODE_ARGENT)
 
290 DR_SoumissionElec.Sections("Section5").Controls("lblGrandTotalAR").Caption = Conversion(rstProjSoum.Fields("prixtotal"), MODE_ARGENT)
 
2  If rstProjSoum.Fields("MontantForfait") <> "" Then
 DR_SoumissionElec.Sections("Section5").Controls("shpCadreForfait").Visible = True
 DR_SoumissionElec.Sections("Section5").Controls("lblTitreForfait").Visible = True
 DR_SoumissionElec.Sections("Section5").Controls("lblForfait").Visible = True

 DR_SoumissionElec.Sections("Section5").Controls("lblTitreForfait").Caption = DR_SoumissionElec.Sections("Section5").Controls("lblTitreForfait").Caption & " ( " & rstProjSoum.Fields("InitialeForfait") & " )"
 DR_SoumissionElec.Sections("Section5").Controls("lblForfait").Caption = rstProjSoum.Fields("MontantForfait")
2  Else
 DR_SoumissionElec.Sections("Section5").Controls("shpCadreForfait").Visible = False
 DR_SoumissionElec.Sections("Section5").Controls("lblTitreForfait").Visible = False
 DR_SoumissionElec.Sections("Section5").Controls("lblForfait").Visible = False
 End If

 '************************************************************************************************
 'SECTION MODIFIÉ PAR GAÉTAN GINGRAS LE   FÉVRIER 2010
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
2   DR_SoumissionElec.Orientation = rptOrientLandscape
 
 Call DR_SoumissionElec.Show(vbModal)
 
2   Call rstImpProjSoum.Close
 Set rstImpProjSoum = Nothing

2   Set rstTemp = Nothing
 
 Screen.MousePointer = vbDefault

29  Exit Sub

Oups:

300 wOups "frmProjSoumElec", "ImprimerProjSoum", Err, Err.number, Err.Description
End Sub

Private Sub ImprimerListePieces(ByVal rstProjSoum As ADODB.Recordset)

 On Error GoTo Oups

 'Impression de la feuille de la liste des pièces
 Dim rstPiece As ADODB.Recordset
 Dim rstTemp As ADODB.Recordset
 Dim rstImpListePiece As ADODB.Recordset
 Dim iCompteurPiece As Integer
 Dim sSousSection As String
 Dim sSection As String
 Dim sNoProjet As String
 Dim sNoSoumission As String
 Dim bAjouterSection As Boolean
 Dim bAjouterSousSection As Boolean
  Dim bAjouterPiece As Boolean

  Set rstPiece = New ADODB.Recordset
  Set rstTemp = New ADODB.Recordset
  Set rstImpListePiece = New ADODB.Recordset

  Call g_connData.Execute("DELETE * FROM Grbimpression_listepiece")

  iCompteurPiece = 1

  Screen.MousePointer = vbHourglass

 'Ouverture du recordset
  If m_eType = TYPE_PROJET Then
sNoProjet = rstProjSoum.Fields("IDProjet")

1 If Not IsNull(rstProjSoum.Fields("IDSoumission")) Then
 sNoSoumission = rstProjSoum.Fields("IDSoumission")
 Else
 sNoSoumission = vbNullString
 End If

 Call rstPiece.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & sNoProjet & "' AND Type = 'E' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
Else
 sNoProjet = vbNullString
 sNoSoumission = rstProjSoum.Fields("IDSoumission")

 Call rstPiece.Open("SELECT * FROM GrbSoumission_Pieces WHERE IDSoumission = '" & sNoSoumission & "' AND Type = 'E' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
End If

1  Do While Not rstPiece.EOF
 If rstPiece.Fields("Visible") = True Then
 bAjouterSection = True
 bAjouterSousSection = True
 bAjouterPiece = True

 rstImpListePiece.CursorLocation = adUseClient

 Call rstImpListePiece.Open("SELECT * FROM GrbImpression_ListePiece WHERE IDSection = '" & rstPiece.Fields("IDSection") & "'", g_connData, adOpenDynamic, adLockOptimistic)

1  If Not rstImpListePiece.EOF Then
 bAjouterSection = False

 Do While Not rstImpListePiece.EOF
 If rstImpListePiece.Fields("NomSousSection") = rstPiece.Fields("SousSection") Then
 bAjouterSousSection = False

 If rstPiece.Fields("NumItem") <> "Texte" And rstPiece.Fields("NumItem") <> "Text" Then
 If rstImpListePiece.Fields("NumItem") = rstPiece.Fields("NumItem") Then
 bAjouterPiece = False

 rstImpListePiece.Fields("Qté") = Replace(CDbl(rstImpListePiece.Fields("Qté")) + CDbl(rstPiece.Fields("Qté")), ".", ",")

 If Not IsNull(rstImpListePiece.Fields("ID")) Then
 If rstImpListePiece.Fields("ID") <> "" Then
 rstImpListePiece.Fields("ID") = rstImpListePiece.Fields("ID") & ", " & rstPiece.Fields("ID")
 Else
 rstImpListePiece.Fields("ID") = rstPiece.Fields("ID")
 End If
 Else
 rstImpListePiece.Fields("ID") = rstPiece.Fields("ID")
 End If

 Call rstImpListePiece.Update

 If rstImpListePiece.Fields("Qté") = 0 Then
 Call rstImpListePiece.Delete

 rstImpListePiece.Filter = "NomSousSection = '" & Replace(rstPiece.Fields("SousSection"), "'", "''") & "'"

 If rstImpListePiece.RecordCount = 1 Then
 Call rstImpListePiece.Delete

 rstImpListePiece.Filter = "IDSection = '" & rstPiece.Fields("IDSection") & "'"

 If rstImpListePiece.RecordCount = 1 Then
 Call rstImpListePiece.Delete
 End If
 End If

 rstImpListePiece.Filter = ""

 End If

 Exit Do
 End If
 Else
 Exit Do
 End If
 End If

 Call rstImpListePiece.MoveNext
 Loop
 End If

 If bAjouterSection = True Then
 If m_eLangage = ANGLAIS Then
4 sSection = "NomSectionEN"
4 Else
4 sSection = "NomSectionFR"
4 End If

4 Call rstTemp.Open("SELECT " & sSection & " FROM GrbSoumProjSectionElec WHERE IDSection = " & rstPiece.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)

 'Ajoute la section dans la liste de pièces
4 Call rstImpListePiece.AddNew

4 rstImpListePiece.Fields("NoLigne") = iCompteurPiece
4 rstImpListePiece.Fields("IDSoumission") = sNoSoumission

4 If Not IsNull(rstTemp.Fields(sSection)) Then
4 rstImpListePiece.Fields("Section") = rstTemp.Fields(sSection)
4 Else
4  rstImpListePiece.Fields("Section") = " "
4  End If

4  rstImpListePiece.Fields("IDSection") = rstPiece.Fields("IDSection")

4  Call rstImpListePiece.Update

4  iCompteurPiece = iCompteurPiece + 1

4  Call rstTemp.Close
4  End If

4  If bAjouterSousSection = True Then
50 sSousSection = rstPiece.Fields("SousSection")

 If sSousSection = S_PAS_SOUS_SECTION Then
 sSousSection = " "
 End If

 Call rstImpListePiece.AddNew

 rstImpListePiece.Fields("NoLigne") = iCompteurPiece
 rstImpListePiece.Fields("IDSoumission") = sNoSoumission
 rstImpListePiece.Fields("SousSection") = sSousSection
 rstImpListePiece.Fields("NomSousSection") = rstPiece.Fields("SousSection")
 rstImpListePiece.Fields("IDSection") = rstPiece.Fields("IDSection")

 Call rstImpListePiece.Update

 iCompteurPiece = iCompteurPiece + 1
5  End If

5  If bAjouterPiece = True Then
 'Ajoute la pièce à la liste de pièces
5  Call rstImpListePiece.AddNew

5  rstImpListePiece.Fields("NoLigne") = iCompteurPiece
5  rstImpListePiece.Fields("IDSoumission") = sNoSoumission
5  rstImpListePiece.Fields("NumItem") = rstPiece.Fields("NumItem")
5  rstImpListePiece.Fields("Qté") = rstPiece.Fields("Qté")

5  If m_eLangage = ANGLAIS Then
60 rstImpListePiece.Fields("Description") = rstPiece.Fields("Desc_EN")
  Else
  rstImpListePiece.Fields("Description") = rstPiece.Fields("Desc_FR")
  End If

  rstImpListePiece.Fields("Manufact") = rstPiece.Fields("Manufact")

  If m_eType = TYPE_PROJET Then
  rstImpListePiece.Fields("ID") = rstPiece.Fields("ID")
  End If

  rstImpListePiece.Fields("IDSection") = rstPiece.Fields("IDSection")
  rstImpListePiece.Fields("NomSousSection") = rstPiece.Fields("SousSection")

  Call rstImpListePiece.Update

  iCompteurPiece = iCompteurPiece + 1
6  End If

6  Call rstImpListePiece.Close
6  End If

 'Prochaine enregistrement
6  Call rstPiece.MoveNext
6  Loop

 ''''''''''''''''''''''''''''''''''''''''''''''''''
 ' Rapport liste pièce, met dans l'ordre de ligne '
 ''''''''''''''''''''''''''''''''''''''''''''''''''
6  rstImpListePiece.CursorLocation = adUseClient

6  Call rstImpListePiece.Open("SELECT * FROM Grbimpression_Listepiece WHERE TRIM(IDSoumission) = '" & Trim$(sNoSoumission) & "' ORDER BY noligne", g_connData, adOpenDynamic, adLockOptimistic)

6  Set DR_Liste_piece.DataSource = rstImpListePiece

70 Call TraduireImpressionListePiece

 'Affiche la date
70 DR_Liste_piece.Sections("Section3").Controls("lblDate").Caption = ConvertDate(Date)

  DR_Liste_piece.Sections("Section4").Controls("lblProjet").Caption = sNoProjet

 'Affiche l 'entête
  DR_Liste_piece.Sections("Section4").Controls("lblSoumission").Caption = sNoSoumission

  DR_Liste_piece.Sections("Section4").Controls("lblDescription").Caption = rstProjSoum.Fields("Description")

  Call rstTemp.Open("SELECT NomClient FROM GrbClient WHERE IDClient = " & rstProjSoum.Fields("IDClient"), g_connData, adOpenDynamic, adLockOptimistic)

  DR_Liste_piece.Sections("Section4").Controls("lblClient").Caption = rstTemp.Fields("NomClient")

  Call rstTemp.Close

  Call rstTemp.Open("SELECT NomContact FROM GrbContact WHERE IDContact = " & rstProjSoum.Fields("IDContact"), g_connData, adOpenDynamic, adLockOptimistic)

  DR_Liste_piece.Sections("Section4").Controls("lblContact").Caption = rstTemp.Fields("nomcontact")

  Call rstTemp.Close

 'Affiche le rapport liste des pieces
  DR_Liste_piece.Orientation = rptOrientPortrait

   Call DR_Liste_piece.Show(vbModal)

   Call rstImpListePiece.Close
7  Set rstImpListePiece = Nothing

7  Set rstTemp = Nothing

7  Screen.MousePointer = vbDefault

7  Exit Sub

Oups:

7  wOups "frmProjSoumElec", "ImprimerListePieces", Err, Err.number, Err.Description
End Sub

Private Sub CalculerTempsReelsImpression(ByVal sNoProjet As String)

 On Error GoTo Oups

 Dim rstTotal As ADODB.Recordset
 Dim sDateDebut As String
 Dim sDateFin As String
 Dim sTotal As String
 Dim sFilterNoProjet As String

 If Right$(sNoProjet, 2) = "99" Then
 sFilterNoProjet = "LEFT(NoProjet, 6) = '" & Left$(sNoProjet, 6) & "'"
 Else
 sFilterNoProjet = "NoProjet = '" & sNoProjet & "'"
 End If

  Set rstTotal = New ADODB.Recordset
 
  sDateDebut = "TIMESERIAL(Left(GrbPunch.HeureDébut,2),RIGHT(GrbPunch.HeureDébut,2),0)"
 
  sDateFin = "TIMESERIAL(Left(GrbPunch.HeureFin,2),RIGHT(GrbPunch.HeureFin,2),0)"
 
  sTotal = "(SUM(" & sDateFin & " - " & sDateDebut & ") * 24) As Total"

  rstTotal.CursorLocation = adUseServer

  Call rstTotal.Open("SELECT Type, " & sTotal & " FROM GrbPunch WHERE " & sFilterNoProjet & " And HeureFin is not null AND HeureDébut is not null GROUP BY Type", g_connData, adOpenDynamic, adLockOptimistic)

  DR_SoumissionElec.Sections("Section5").Controls("lblTempsDessinReel").Caption = "0"
  DR_SoumissionElec.Sections("Section5").Controls("lblTempsFabricationReel").Caption = "0"
10 DR_SoumissionElec.Sections("Section5").Controls("lblTempsAssemblageReel").Caption = "0"
DR_SoumissionElec.Sections("Section5").Controls("lblTempsProgInterfaceReel").Caption = "0"
DR_SoumissionElec.Sections("Section5").Controls("lblTempsProgAutomateReel").Caption = "0"
DR_SoumissionElec.Sections("Section5").Controls("lblTempsProgRobotReel").Caption = "0"
DR_SoumissionElec.Sections("Section5").Controls("lblTempsVisionReel").Caption = "0"
DR_SoumissionElec.Sections("Section5").Controls("lblTempsTestReel").Caption = "0"
DR_SoumissionElec.Sections("Section5").Controls("lblTempsInstallationReel").Caption = "0"
DR_SoumissionElec.Sections("Section5").Controls("lblTempsMiseServiceReel").Caption = "0"
DR_SoumissionElec.Sections("Section5").Controls("lblTempsFormationReel").Caption = "0"
DR_SoumissionElec.Sections("Section5").Controls("lblTempsGestionReel").Caption = "0"
DR_SoumissionElec.Sections("Section5").Controls("lblTempsShippingReel").Caption = "0"

Do While Not rstTotal.EOF
If Not IsNull(rstTotal.Fields("Total")) Then
 Select Case rstTotal.Fields("Type")
 Case "Dessin": DR_SoumissionElec.Sections("Section5").Controls("lblTempsDessinReel").Caption = Round(rstTotal.Fields("Total"), 2)
 Case "Fabrication": DR_SoumissionElec.Sections("Section5").Controls("lblTempsFabricationReel").Caption = Round(rstTotal.Fields("Total"), 2)
 Case "Assemblage": DR_SoumissionElec.Sections("Section5").Controls("lblTempsAssemblageReel").Caption = Round(rstTotal.Fields("Total"), 2)
 Case "ProgInterface": DR_SoumissionElec.Sections("Section5").Controls("lblTempsProgInterfaceReel").Caption = Round(rstTotal.Fields("Total"), 2)
 Case "ProgAutomate": DR_SoumissionElec.Sections("Section5").Controls("lblTempsProgAutomateReel").Caption = Round(rstTotal.Fields("Total"), 2)
 Case "ProgRobot": DR_SoumissionElec.Sections("Section5").Controls("lblTempsProgRobotReel").Caption = Round(rstTotal.Fields("Total"), 2)
1  Case "Vision": DR_SoumissionElec.Sections("Section5").Controls("lblTempsVisionReel").Caption = Round(rstTotal.Fields("Total"), 2)
 Case "Test": DR_SoumissionElec.Sections("Section5").Controls("lblTempsTestReel").Caption = Round(rstTotal.Fields("Total"), 2)
 Case "Installation": DR_SoumissionElec.Sections("Section5").Controls("lblTempsInstallationReel").Caption = Round(rstTotal.Fields("Total"), 2)
 Case "MiseService": DR_SoumissionElec.Sections("Section5").Controls("lblTempsMiseServiceReel").Caption = Round(rstTotal.Fields("Total"), 2)
 Case "Formation": DR_SoumissionElec.Sections("Section5").Controls("lblTempsFormationReel").Caption = Round(rstTotal.Fields("Total"), 2)
 Case "Gestion": DR_SoumissionElec.Sections("Section5").Controls("lblTempsGestionReel").Caption = Round(rstTotal.Fields("Total"), 2)
 Case "Shipping": DR_SoumissionElec.Sections("Section5").Controls("lblTempsShippingReel").Caption = Round(rstTotal.Fields("Total"), 2)
 End Select
 End If

 Call rstTotal.MoveNext
Loop

Call rstTotal.Close
 
Call rstTotal.Open("SELECT " & sTotal & " FROM GrbPunch WHERE " & sFilterNoProjet & " And HeureFin is not null AND HeureDébut is not null", g_connData, adOpenDynamic, adLockOptimistic)

2  If Not IsNull(rstTotal.Fields("Total")) Then
 DR_SoumissionElec.Sections("Section5").Controls("lblTotalTempsRHReel").Caption = Round(rstTotal.Fields("Total"), 2)
2  Else
 DR_SoumissionElec.Sections("Section5").Controls("lblTotalTempsRHReel").Caption = "0"
2  End If

Call rstTotal.Close
2  Set rstTotal = Nothing

Exit Sub

Oups:

30 wOups "frmProjSoumElec", "CalculerTempsReels", Err, Err.number, Err.Description
End Sub

Private Sub ImprimerProjSoumFacturation(ByVal rstProjSoum As ADODB.Recordset, ByVal sNoFacture As String)

 On Error GoTo Oups

 'Impression de la feuille de soumission
 Dim rstPiece As ADODB.Recordset
 Dim rstTemp As ADODB.Recordset
 Dim rstImpProjSoum As ADODB.Recordset
 Dim sOrdreSection As String
 Dim iCompteurSoum As Integer
 Dim sSousSection As String
 Dim sSousSectionRS As String
 Dim sSection As String
 Dim sNoProjet As String
 Dim sNoSoumission As String
  Dim sCommission As String
  Dim sPrixTotal As String
  Dim sProfit As String
  Dim sTempsFabrication As String
  Dim sTotalPiece As String
  Dim sImprevue As String
  Dim sTotalTemps As String
  Dim sManuel As String
10 Dim dblTotalTemps As Double
Dim dblTempsDessin As Double
Dim dblTempsFabrication As Double
Dim dblTempsAssemblage As Double
Dim dblTempsProgInterface As Double
Dim dblTempsProgAutomate As Double
Dim dblTempsProgRobot As Double
Dim dblTempsVision As Double
Dim dblTempsTest As Double
Dim dblTempsInstallation As Double
Dim dblTempsMiseService As Double
Dim dblTempsFormation As Double
1  Dim dblTempsGestion As Double
Dim dblTempsShipping As Double
 Dim dblTotalHebergement As Double
Dim dblTotalRepas As Double
 Dim dblTotalTransport As Double
Dim dblTotalUniteMobile As Double
 Dim dblPrixEmballage As Double
1  Dim dblTotalReste As Double
 Dim dblTotalAutre As Double

 Set rstPiece = New ADODB.Recordset
Set rstTemp = New ADODB.Recordset
Set rstImpProjSoum = New ADODB.Recordset

 'Supprime les données de l'impression
Call g_connData.Execute("DELETE * FROM Grbimpression_soumission")

iCompteurSoum = 1
 
Screen.MousePointer = vbHourglass

Call rstImpProjSoum.Open("SELECT * FROM Grbimpression_soumission", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Ouverture du recordset
sNoProjet = rstProjSoum.Fields("IDProjet")
sNoSoumission = rstProjSoum.Fields("IDSoumission")

Call rstPiece.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & sNoProjet & "' AND Type = 'E' AND Facturation = '" & sNoFacture & "' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
 
Do While Not rstPiece.EOF
sSousSectionRS = rstPiece.Fields("SousSection")
 
 If sSousSectionRS = S_PAS_SOUS_SECTION Then
 sSousSectionRS = " "
 End If
 
If sOrdreSection <> rstPiece.Fields("OrdreSection") Then
 'Remplis la table impression_soumission
 'Ajoute seulement la section
 sOrdreSection = rstPiece.Fields("OrdreSection")
 
 If m_eLangage = ANGLAIS Then
 sSection = "NomSectionEN"
 Else
 sSection = "NomSectionFR"
 End If
 
 Call rstTemp.Open("SELECT " & sSection & " FROM GrbSoumProjSectionElec WHERE IDSection = " & rstPiece.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
 
 'Ajoute la section dans la soumission
 Call rstImpProjSoum.AddNew
 
 rstImpProjSoum.Fields("NoLigne") = iCompteurSoum

 rstImpProjSoum.Fields("IDSoumission") = sNoProjet
 
 If Not IsNull(rstTemp.Fields(sSection)) Then
 rstImpProjSoum.Fields("NomSection") = rstTemp.Fields(sSection)
 Else
 rstImpProjSoum.Fields("NomSection") = " "
 End If
 
 Call rstImpProjSoum.Update
 
 iCompteurSoum = iCompteurSoum + 1
 
 Call rstTemp.Close
 
 sSousSection = rstPiece.Fields("SousSection")
 
 If sSousSection = S_PAS_SOUS_SECTION Then
 sSousSection = " "
 End If
 
 Call rstImpProjSoum.AddNew
 
 rstImpProjSoum.Fields("NoLigne") = iCompteurSoum

4 If m_eType = TYPE_PROJET Then
4 rstImpProjSoum.Fields("IDSoumission") = sNoProjet
4 Else
4 rstImpProjSoum.Fields("IDSoumission") = sNoSoumission
4 End If

4 rstImpProjSoum.Fields("SousSection") = sSousSection
 
4 Call rstImpProjSoum.Update

4 iCompteurSoum = iCompteurSoum + 1
4 Else
 'Ajoute une soussection dans impression_soum
4 If sSousSection <> sSousSectionRS Then
4 sSousSection = sSousSectionRS
 
4  Call rstImpProjSoum.AddNew
 
4  rstImpProjSoum.Fields("NoLigne") = iCompteurSoum

4  rstImpProjSoum.Fields("IDSoumission") = sNoProjet

4  rstImpProjSoum.Fields("SousSection") = sSousSectionRS
 
4  Call rstImpProjSoum.Update
 
4  iCompteurSoum = iCompteurSoum + 1
4  End If
4  End If
 
 'Ajoute une piece dans impression_soum
50 Call rstImpProjSoum.AddNew
 
5 rstImpProjSoum.Fields("NoLigne") = iCompteurSoum

 rstImpProjSoum.Fields("IDSoumission") = sNoProjet

 rstImpProjSoum.Fields("NumItem") = rstPiece.Fields("NumItem")
 rstImpProjSoum.Fields("Qté") = rstPiece.Fields("Qté")
 
 If m_eLangage = ANGLAIS Then
 rstImpProjSoum.Fields("Description") = rstPiece.Fields("DESC_EN")
 Else
 rstImpProjSoum.Fields("Description") = rstPiece.Fields("DESC_FR")
 End If

 '************************************************************************************************
 'SECTION MODIFIÉ PAR GAÉTAN GINGRAS LE 0  FÉVRIER 2010
 '************************************************************************************************
 rstImpProjSoum.Fields("MANUFACT") = rstPiece.Fields("MANUFACT")
 'rstImpProjSoum.Fields("PRIX_LIST") = rstPiece.Fields("PRIX_LIST")

5  'If Not IsNull(rstPiece.Fields("ESCOMPTE")) And Trim(rstPiece.Fields("ESCOMPTE")) <> vbNullString Then
5  ' If rstPiece.Fields("ESCOMPTE") > 0 Then
5  ' rstImpProjSoum.Fields("ESCOMPTE") = Conversion(Replace(rstPiece.Fields("escompte"), ".", ","), MODE_POURCENT)
5  ' Else
5  ' rstImpProjSoum.Fields("ESCOMPTE") = Conversion(Replace(rstPiece.Fields("escompte"), ".", ",") * 100, MODE_POURCENT)
5  ' End If
5  'Else
5  ' rstImpProjSoum.Fields("ESCOMPTE") = Conversion(0, MODE_POURCENT)
60 'End If
 
  rstImpProjSoum.Fields("PRIX_NET") = rstPiece.Fields("PRIX_NET")
 
  Call rstTemp.Open("SELECT NomFournisseur FROM GrbFournisseur WHERE IDFRS = " & rstPiece.Fields("IDFRS"), g_connData, adOpenDynamic, adLockOptimistic)
 
  If Not rstTemp.EOF Then
  rstImpProjSoum.Fields("NomFournisseur") = rstTemp.Fields("NomFournisseur")
  End If
 
  Call rstTemp.Close
 
  'rstImpProjSoum.Fields("TEMPS") = rstPiece.Fields("TEMPS")
  'rstImpProjSoum.Fields("TEMPS_TOTAL") = rstPiece.Fields("TEMPS_TOTAL")
  rstImpProjSoum.Fields("PRIX_TOTAL") = rstPiece.Fields("PRIX_TOTAL")
  rstImpProjSoum.Fields("PROFIT_ARGENT") = rstPiece.Fields("PROFIT_ARGENT")

 'AJOUT DE CETTE SECTION PAR GAÉTAN GINGRAS LE 0  FÉVRIER 2010
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
 Else 'il n'y a pas de champ date de réception et de commande dans la table GrbSoumission_Pièces
 rstImpProjSoum.Fields("DateReception") = ""
 rstImpProjSoum.Fields("DateCommande") = ""
 End If
 '************************************************************************************************
 'FIN DE LA SECTION DE MODIFICATION
 '************************************************************************************************
 
  Call rstImpProjSoum.Update
 
6  iCompteurSoum = iCompteurSoum + 1
 
 'Prochaine enreg
6  Call rstPiece.MoveNext
6  Loop
 
 'Ferme les tables
6  Call rstImpProjSoum.Close
 
 '''''''''''''''''''''''''''''''''''''''''''''''''
 ' Rapport soumission, met dans l'ordre de ligne '
 '''''''''''''''''''''''''''''''''''''''''''''''''
 
6  Call rstImpProjSoum.Open("SELECT * FROM Grbimpression_soumission WHERE IDSoumission = '" & sNoProjet & "' ORDER BY NoLigne", g_connData, adOpenDynamic, adLockOptimistic)
 
6  Set DR_SoumissionElec.DataSource = rstImpProjSoum

6  Call CalculerPrixFacturation(sNoFacture, sCommission, sPrixTotal, sProfit, sTempsFabrication, sTotalPiece, sImprevue, sTotalTemps, sManuel)
 
6  Call TraduireImpressionSoumission
 
70 DR_SoumissionElec.Sections("Section5").Controls("shpCadrePrixReception").Visible = False
70 DR_SoumissionElec.Sections("Section5").Controls("lblTitrePrixReception").Visible = False
  DR_SoumissionElec.Sections("Section5").Controls("lblPrixReception").Visible = False

  DR_SoumissionElec.Sections("Section5").Controls("shpCadrePrixSoumission").Visible = False
  DR_SoumissionElec.Sections("Section5").Controls("lblTitrePrixSoumission").Visible = False
  DR_SoumissionElec.Sections("Section5").Controls("lblPrixSoumission").Visible = False

  DR_SoumissionElec.Sections("Section5").Controls("shpCadreForfait").Visible = False
  DR_SoumissionElec.Sections("Section5").Controls("lblTitreForfait").Visible = False
  DR_SoumissionElec.Sections("Section5").Controls("lblForfait").Visible = False

  DR_SoumissionElec.Sections("Section2").Controls("lblTitreNoFacture").Visible = True
  DR_SoumissionElec.Sections("Section2").Controls("lblNoFacture").Visible = True

  DR_SoumissionElec.Sections("Section2").Controls("lblNoFacture").Caption = sNoFacture
 
 'Affiche la date
   DR_SoumissionElec.Sections("Section3").Controls("lbldate").Caption = ConvertDate(Date)
 
 'Affiche entete
   DR_SoumissionElec.Sections("Section2").Controls("lblSoumission").Caption = sNoSoumission
 
7  DR_SoumissionElec.Sections("Section2").Controls("lblprojet").Caption = sNoProjet
 
7  DR_SoumissionElec.Sections("Section2").Controls("lbldescription").Caption = rstProjSoum.Fields("Description")
 
7  Call rstTemp.Open("SELECT NomClient FROM GrbClient WHERE IDClient = " & rstProjSoum.Fields("IDClient"), g_connData, adOpenDynamic, adLockOptimistic)
 
7  DR_SoumissionElec.Sections("Section2").Controls("lblclient").Caption = rstTemp.Fields("NomClient")
 
7  Call rstTemp.Close
 
7  Call rstTemp.Open("SELECT NomContact FROM GrbContact WHERE IDContact = " & rstProjSoum.Fields("IDContact"), g_connData, adOpenDynamic, adLockOptimistic)
 
80 DR_SoumissionElec.Sections("Section2").Controls("lblcontact").Caption = rstTemp.Fields("NomContact")
 
80 Call rstTemp.Close
 
 'affiche pied d'etat
  If Not IsNull(rstProjSoum.Fields("TauxDessin")) Then
  DR_SoumissionElec.Sections("Section5").Controls("lblTauxDessin").Caption = rstProjSoum.Fields("TauxDessin")
  Else
  DR_SoumissionElec.Sections("Section5").Controls("lblTauxDessin").Caption = "0"
  End If

  If Not IsNull(rstProjSoum.Fields("TauxFabrication")) Then
  DR_SoumissionElec.Sections("Section5").Controls("lblTauxFabrication").Caption = rstProjSoum.Fields("TauxFabrication")
  Else
  DR_SoumissionElec.Sections("Section5").Controls("lblTauxFabrication").Caption = "0"
  End If

   If Not IsNull(rstProjSoum.Fields("TauxAssemblage")) Then
   DR_SoumissionElec.Sections("Section5").Controls("lblTauxAssemblage").Caption = rstProjSoum.Fields("TauxAssemblage")
   Else
   DR_SoumissionElec.Sections("Section5").Controls("lblTauxAssemblage").Caption = "0"
8  End If

8  If Not IsNull(rstProjSoum.Fields("TauxProgInterface")) Then
8  DR_SoumissionElec.Sections("Section5").Controls("lblTauxProgInterface").Caption = rstProjSoum.Fields("TauxProgInterface")
8  Else
90 DR_SoumissionElec.Sections("Section5").Controls("lblTauxProgInterface").Caption = "0"
90 End If

  If Not IsNull(rstProjSoum.Fields("TauxProgAutomate")) Then
  DR_SoumissionElec.Sections("Section5").Controls("lblTauxProgAutomate").Caption = rstProjSoum.Fields("TauxProgAutomate")
  Else
  DR_SoumissionElec.Sections("Section5").Controls("lblTauxProgAutomate").Caption = "0"
  End If

  If Not IsNull(rstProjSoum.Fields("TauxProgRobot")) Then
  DR_SoumissionElec.Sections("Section5").Controls("lblTauxProgRobot").Caption = rstProjSoum.Fields("TauxProgRobot")
  Else
  DR_SoumissionElec.Sections("Section5").Controls("lblTauxProgRobot").Caption = "0"
  End If

 If Not IsNull(rstProjSoum.Fields("TauxVision")) Then
   DR_SoumissionElec.Sections("Section5").Controls("lblTauxVision").Caption = rstProjSoum.Fields("TauxVision")
 Else
   DR_SoumissionElec.Sections("Section5").Controls("lblTauxVision").Caption = "0"
 End If

   If Not IsNull(rstProjSoum.Fields("TauxTest")) Then
 DR_SoumissionElec.Sections("Section5").Controls("lblTauxTest").Caption = rstProjSoum.Fields("TauxTest")
9  Else
 DR_SoumissionElec.Sections("Section5").Controls("lblTauxTest").Caption = "0"
100 End If

10 If Not IsNull(rstProjSoum.Fields("TauxInstallation")) Then
1DR_SoumissionElec.Sections("Section5").Controls("lblTauxInstallation").Caption = rstProjSoum.Fields("TauxInstallation")
Else
1DR_SoumissionElec.Sections("Section5").Controls("lblTauxInstallation").Caption = "0"
End If

10 If Not IsNull(rstProjSoum.Fields("TauxMiseService")) Then
 DR_SoumissionElec.Sections("Section5").Controls("lblTauxMiseService").Caption = rstProjSoum.Fields("TauxMiseService")
10 Else
 DR_SoumissionElec.Sections("Section5").Controls("lblTauxMiseService").Caption = "0"
10 End If

10  If Not IsNull(rstProjSoum.Fields("TauxFormation")) Then
10  DR_SoumissionElec.Sections("Section5").Controls("lblTauxFormation").Caption = rstProjSoum.Fields("TauxFormation")
107Else
10  DR_SoumissionElec.Sections("Section5").Controls("lblTauxFormation").Caption = "0"
108End If

10  If Not IsNull(rstProjSoum.Fields("TauxGestion")) Then
10  DR_SoumissionElec.Sections("Section5").Controls("lblTauxGestion").Caption = rstProjSoum.Fields("TauxGestion")
10  Else
110 DR_SoumissionElec.Sections("Section5").Controls("lblTauxGestion").Caption = "0"
110 End If

11 If Not IsNull(rstProjSoum.Fields("TauxShipping")) Then
1 DR_SoumissionElec.Sections("Section5").Controls("lblTauxShipping").Caption = rstProjSoum.Fields("TauxShipping")
11 Else
1 DR_SoumissionElec.Sections("Section5").Controls("lblTauxShipping").Caption = "0"
11 End If

11 If Not IsNull(rstProjSoum.Fields("TempsDessin")) Then
1 DR_SoumissionElec.Sections("Section5").Controls("lblTempsDessinSoum").Caption = rstProjSoum.Fields("TempsDessin")
11 Else
1 DR_SoumissionElec.Sections("Section5").Controls("lblTempsDessinSoum").Caption = "0"
11 End If

116DR_SoumissionElec.Sections("Section5").Controls("lblTempsFabricationSoum").Caption = sTempsFabrication

11  If Not IsNull(rstProjSoum.Fields("TempsAssemblage")) Then
 DR_SoumissionElec.Sections("Section5").Controls("lblTempsAssemblageSoum").Caption = rstProjSoum.Fields("TempsAssemblage")
11  Else
 DR_SoumissionElec.Sections("Section5").Controls("lblTempsAssemblageSoum").Caption = "0"
11  End If

1 If Not IsNull(rstProjSoum.Fields("TempsProgInterface")) Then
11  DR_SoumissionElec.Sections("Section5").Controls("lblTempsProgInterfaceSoum").Caption = rstProjSoum.Fields("TempsProgInterface")
1 Else
1 DR_SoumissionElec.Sections("Section5").Controls("lblTempsProgInterfaceSoum").Caption = "0"
12 End If

12 If Not IsNull(rstProjSoum.Fields("TempsProgAutomate")) Then
1 DR_SoumissionElec.Sections("Section5").Controls("lblTempsProgAutomateSoum").Caption = rstProjSoum.Fields("TempsProgAutomate")
12 Else
1 DR_SoumissionElec.Sections("Section5").Controls("lblTempsProgAutomateSoum").Caption = "0"
12 End If

12 If Not IsNull(rstProjSoum.Fields("TempsProgRobot")) Then
1 DR_SoumissionElec.Sections("Section5").Controls("lblTempsProgRobotSoum").Caption = rstProjSoum.Fields("TempsProgRobot")
12 Else
1 DR_SoumissionElec.Sections("Section5").Controls("lblTempsProgRobotSoum").Caption = "0"
126End If

12  If Not IsNull(rstProjSoum.Fields("TempsVision")) Then
12  DR_SoumissionElec.Sections("Section5").Controls("lblTempsVisionSoum").Caption = rstProjSoum.Fields("TempsVision")
12  Else
12  DR_SoumissionElec.Sections("Section5").Controls("lblTempsVisionSoum").Caption = "0"
12  End If

12  If Not IsNull(rstProjSoum.Fields("TempsTest")) Then
1 DR_SoumissionElec.Sections("Section5").Controls("lblTempsTestSoum").Caption = rstProjSoum.Fields("TempsTest")
130Else
13DR_SoumissionElec.Sections("Section5").Controls("lblTempsTestSoum").Caption = "0"
13 End If

13 If Not IsNull(rstProjSoum.Fields("TempsInstallation")) Then
1 DR_SoumissionElec.Sections("Section5").Controls("lblTempsInstallationSoum").Caption = rstProjSoum.Fields("TempsInstallation")
13 Else
1 DR_SoumissionElec.Sections("Section5").Controls("lblTempsInstallationSoum").Caption = "0"
13 End If

13 If Not IsNull(rstProjSoum.Fields("TempsMiseService")) Then
1 DR_SoumissionElec.Sections("Section5").Controls("lblTempsMiseServiceSoum").Caption = rstProjSoum.Fields("TempsMiseService")
13 Else
1 DR_SoumissionElec.Sections("Section5").Controls("lblTempsMiseServiceSoum").Caption = "0"
136End If

13  If Not IsNull(rstProjSoum.Fields("TempsFormation")) Then
13  DR_SoumissionElec.Sections("Section5").Controls("lblTempsFormationSoum").Caption = rstProjSoum.Fields("TempsFormation")
13  Else
1DR_SoumissionElec.Sections("Section5").Controls("lblTempsFormationSoum").Caption = "0"
1End If

13  If Not IsNull(rstProjSoum.Fields("TempsGestion")) Then
1 DR_SoumissionElec.Sections("Section5").Controls("lblTempsGestionSoum").Caption = rstProjSoum.Fields("TempsGestion")
140Else
14DR_SoumissionElec.Sections("Section5").Controls("lblTempsGestionSoum").Caption = "0"
14 End If

14 If Not IsNull(rstProjSoum.Fields("TempsShipping")) Then
14 DR_SoumissionElec.Sections("Section5").Controls("lblTempsShippingSoum").Caption = rstProjSoum.Fields("TempsShipping")
14 Else
14 DR_SoumissionElec.Sections("Section5").Controls("lblTempsShippingSoum").Caption = "0"
14 End If

14 If IsNumeric(rstProjSoum.Fields("TempsDessin")) Then
14 dblTempsDessin = CDbl(rstProjSoum.Fields("TempsDessin"))
14 Else
14 dblTempsDessin = 0
146End If

14  If IsNumeric(sTempsFabrication) Then
14  dblTempsFabrication = CDbl(sTempsFabrication)
14  Else
14  dblTempsFabrication = 0
14  End If

14  If IsNumeric(rstProjSoum.Fields("TempsAssemblage")) Then
14  dblTempsAssemblage = CDbl(rstProjSoum.Fields("TempsAssemblage"))
150Else
15dblTempsAssemblage = 0
End If

1 If IsNumeric(rstProjSoum.Fields("TempsProgInterface")) Then
 dblTempsProgInterface = CDbl(rstProjSoum.Fields("TempsProgInterface"))
Else
 dblTempsProgInterface = 0
End If

If IsNumeric(rstProjSoum.Fields("TempsProgAutomate")) Then
 dblTempsProgAutomate = CDbl(rstProjSoum.Fields("TempsProgAutomate"))
Else
 dblTempsProgAutomate = 0
156End If

15  If IsNumeric(rstProjSoum.Fields("TempsProgRobot")) Then
15  dblTempsProgRobot = CDbl(rstProjSoum.Fields("TempsProgRobot"))
15  Else
15  dblTempsProgRobot = 0
15  End If

15  If IsNumeric(rstProjSoum.Fields("TempsVision")) Then
15  dblTempsVision = CDbl(rstProjSoum.Fields("TempsVision"))
160Else
16dblTempsVision = 0
1  End If

1  If IsNumeric(rstProjSoum.Fields("TempsTest")) Then
 dblTempsTest = CDbl(rstProjSoum.Fields("TempsTest"))
1  Else
 dblTempsTest = 0
1  End If

1  If IsNumeric(rstProjSoum.Fields("TempsInstallation")) Then
 dblTempsInstallation = CDbl(rstProjSoum.Fields("TempsInstallation"))
1  Else
 dblTempsInstallation = 0
166End If

16  If IsNumeric(rstProjSoum.Fields("TempsMiseService")) Then
16  dblTempsMiseService = CDbl(rstProjSoum.Fields("TempsMiseService"))
16  Else
16  dblTempsMiseService = 0
16  End If

16  If IsNumeric(rstProjSoum.Fields("TempsFormation")) Then
16  dblTempsFormation = CDbl(rstProjSoum.Fields("TempsFormation"))
170Else
 dblTempsFormation = 0
1  End If

1  If IsNumeric(rstProjSoum.Fields("TempsGestion")) Then
 dblTempsGestion = CDbl(rstProjSoum.Fields("TempsGestion"))
1  Else
 dblTempsGestion = 0
1  End If

1  If IsNumeric(rstProjSoum.Fields("TempsShipping")) Then
 dblTempsShipping = CDbl(rstProjSoum.Fields("TempsShipping"))
1  Else
 dblTempsShipping = 0
176End If

1   dblTotalTemps = dblTempsDessin + _
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
 
177DR_SoumissionElec.Sections("Section5").Controls("lblTotalTempsRHSoum").Caption = dblTotalTemps

17  Call CalculerTempsReelsImpression(rstProjSoum.Fields("IDProjet"))

 'Autres frais
17  If Not IsNull(rstProjSoum.Fields("NbrePersonne")) Then
17  DR_SoumissionElec.Sections("Section5").Controls("lblNbrePersonne").Caption = rstProjSoum.Fields("NbrePersonne")
179Else
17  DR_SoumissionElec.Sections("Section5").Controls("lblNbrePersonne").Caption = "0"
180End If

If Not IsNull(rstProjSoum.Fields("TempsHebergement")) Then
 DR_SoumissionElec.Sections("Section5").Controls("lblTempsHebergement").Caption = rstProjSoum.Fields("TempsHebergement")
1  Else
 DR_SoumissionElec.Sections("Section5").Controls("lblTempsHebergement").Caption = "0"
1  End If

1  If Not IsNull(rstProjSoum.Fields("TauxHebergement1")) Then
 DR_SoumissionElec.Sections("Section5").Controls("lblTauxHebergement1").Caption = rstProjSoum.Fields("TauxHebergement1")
1  Else
 DR_SoumissionElec.Sections("Section5").Controls("lblTauxHebergement1").Caption = "0"
1  End If

1  If Not IsNull(rstProjSoum.Fields("TauxHebergement2")) Then
1   DR_SoumissionElec.Sections("Section5").Controls("lblTauxHebergement2").Caption = rstProjSoum.Fields("TauxHebergement2")
1   Else
1   DR_SoumissionElec.Sections("Section5").Controls("lblTauxHebergement2").Caption = "0"
1   End If

18  If Not IsNull(rstProjSoum.Fields("TempsRepas")) Then
18  DR_SoumissionElec.Sections("Section5").Controls("lblTempsRepas").Caption = rstProjSoum.Fields("TempsRepas")
189Else
18  DR_SoumissionElec.Sections("Section5").Controls("lblTempsRepas").Caption = "0"
190End If

If Not IsNull(rstProjSoum.Fields("TauxRepas")) Then
1  DR_SoumissionElec.Sections("Section5").Controls("lblTauxRepas").Caption = rstProjSoum.Fields("TauxRepas")
1  Else
1  DR_SoumissionElec.Sections("Section5").Controls("lblTauxRepas").Caption = "0"
1  End If

1  If Not IsNull(rstProjSoum.Fields("TempsTransport")) Then
1  DR_SoumissionElec.Sections("Section5").Controls("lblTempsTransport").Caption = rstProjSoum.Fields("TempsTransport")
1  Else
1  DR_SoumissionElec.Sections("Section5").Controls("lblTempsTransport").Caption = "0"
1  End If

1  If Not IsNull(rstProjSoum.Fields("TauxTransport")) Then
 DR_SoumissionElec.Sections("Section5").Controls("lblTauxTransport").Caption = rstProjSoum.Fields("TauxTransport")
1   Else
 DR_SoumissionElec.Sections("Section5").Controls("lblTauxTransport").Caption = "0"
1   End If

 If Not IsNull(rstProjSoum.Fields("TempsUniteMobile")) Then
1   DR_SoumissionElec.Sections("Section5").Controls("lblTempsUniteMobile").Caption = rstProjSoum.Fields("TempsUniteMobile")
 Else
19  DR_SoumissionElec.Sections("Section5").Controls("lblTempsUniteMobile").Caption = "0"
200End If

If Not IsNull(rstProjSoum.Fields("TauxUniteMobile")) Then
 DR_SoumissionElec.Sections("Section5").Controls("lblTauxUniteMobile").Caption = rstProjSoum.Fields("TauxUniteMobile")
Else
 DR_SoumissionElec.Sections("Section5").Controls("lblTauxUniteMobile").Caption = "0"
End If

If Not IsNull(rstProjSoum.Fields("PrixEmballage")) Then
 DR_SoumissionElec.Sections("Section5").Controls("lblPrixEmballage").Caption = rstProjSoum.Fields("PrixEmballage")
Else
 DR_SoumissionElec.Sections("Section5").Controls("lblPrixEmballage").Caption = "0"
End If

DR_SoumissionElec.Sections("Section5").Controls("lblPrixManuel").Caption = rstProjSoum.Fields("Total_Manuel")

206DR_SoumissionElec.Sections("Section5").Controls("lblTotalPieceAR").Caption = Conversion(sTotalPiece, MODE_ARGENT)
20  DR_SoumissionElec.Sections("Section5").Controls("lblImprevue").Caption = Conversion(rstProjSoum.Fields("imprevue"), MODE_POURCENT)
207DR_SoumissionElec.Sections("Section5").Controls("lblImprevueAR").Caption = Conversion(sImprevue, MODE_ARGENT)
20  DR_SoumissionElec.Sections("Section5").Controls("lblTotalTempsAR").Caption = Conversion(sTotalTemps, MODE_ARGENT)
208DR_SoumissionElec.Sections("Section5").Controls("lblCommission").Caption = Conversion(rstProjSoum.Fields("commission"), MODE_POURCENT)
20  DR_SoumissionElec.Sections("Section5").Controls("lblCommissionAR").Caption = Conversion(sCommission, MODE_ARGENT)
209DR_SoumissionElec.Sections("Section5").Controls("lblGrandTotalAR").Caption = Conversion(sPrixTotal, MODE_ARGENT)
20  DR_SoumissionElec.Sections("Section5").Controls("lblProfit").Caption = Conversion(rstProjSoum.Fields("profit") * 100, MODE_POURCENT)
210DR_SoumissionElec.Sections("Section5").Controls("lblTotalProfit").Caption = Conversion(sProfit, MODE_ARGENT)

210 dblTotalHebergement = CDbl(rstProjSoum.Fields("TotalHebergement"))

21 dblTotalRepas = CDbl(rstProjSoum.Fields("TotalRepas"))

21 If Not IsNull(rstProjSoum.Fields("TempsTransport")) And Not IsNull(rstProjSoum.Fields("TauxTransport")) Then
2 dblTotalTransport = CDbl(rstProjSoum.Fields("TempsTransport")) * CDbl(rstProjSoum.Fields("TauxTransport"))
21 Else
2 dblTotalTransport = 0
21 End If

21 If Not IsNull(rstProjSoum.Fields("TempsUniteMobile")) And Not IsNull(rstProjSoum.Fields("TauxUniteMobile")) Then
2 dblTotalUniteMobile = CDbl(rstProjSoum.Fields("TempsUniteMobile")) * CDbl(rstProjSoum.Fields("TauxUniteMobile"))
21 Else
2 dblTotalUniteMobile = 0
216End If

21  If Not IsNull(rstProjSoum.Fields("PrixEmballage")) Then
 dblPrixEmballage = CDbl(Replace(rstProjSoum.Fields("PrixEmballage"), ".", ","))
21  Else
 dblPrixEmballage = 0
21  End If
 
2 dblTotalReste = dblTotalHebergement + dblTotalRepas + dblTotalTransport + dblTotalUniteMobile + dblPrixEmballage

21  dblTotalAutre = dblTotalReste + CDbl(rstProjSoum.Fields("total_manuel"))
 
2 DR_SoumissionElec.Sections("Section5").Controls("lblAutre").Caption = Conversion(CStr(dblTotalAutre), MODE_ARGENT)

 '************************************************************************************************
 'SECTION MODIFIÉ PAR GAÉTAN GINGRAS LE 0  FÉVRIER 2010
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
2DR_SoumissionElec.Orientation = rptOrientLandscape
 
22 Call DR_SoumissionElec.Show(vbModal)
 
22 Set rstTemp = Nothing
 
22 Screen.MousePointer = vbDefault

22 Exit Sub

Oups:

22 wOups "frmProjSoumElec", "ImprimerProjSoum", Err, Err.number, Err.Description
End Sub

Private Sub ImprimerListePiecesFacturation(ByVal rstProjSoum As ADODB.Recordset, ByVal sNoFacture As String)
 
 On Error GoTo Oups

 'Impression de la feuille de la liste des pièces
 Dim rstPiece As ADODB.Recordset
 Dim rstTemp As ADODB.Recordset
 Dim rstImpListePiece As ADODB.Recordset
 Dim sOrdreSection As String
 Dim iCompteurPiece As Integer
 Dim sSousSection As String
 Dim sSousSectionRS As String
 Dim sSection As String
 Dim sNoProjet As String
 Dim sNoSoumission As String

  Set rstPiece = New ADODB.Recordset
  Set rstTemp = New ADODB.Recordset
  Set rstImpListePiece = New ADODB.Recordset

  Call g_connData.Execute("DELETE * FROM Grbimpression_listepiece")

  iCompteurPiece = 1
 
  Screen.MousePointer = vbHourglass
 
  Call rstImpListePiece.Open("SELECT * FROM Grbimpression_listepiece", g_connData, adOpenDynamic, adLockOptimistic)

  sOrdreSection = vbNullString

 'Ouverture du recordset
10 sNoProjet = rstProjSoum.Fields("IDProjet")
sNoSoumission = rstProjSoum.Fields("IDSoumission")

Call rstPiece.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & sNoProjet & "' AND Type = 'E' AND Facturation = '" & sNoFacture & "' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
 
Do While Not rstPiece.EOF
 If rstPiece.Fields("Visible") = True Then
 sSousSectionRS = rstPiece.Fields("SousSection")
 
 If sSousSectionRS = S_PAS_SOUS_SECTION Then
 sSousSectionRS = " "
 End If
 
 If sOrdreSection <> rstPiece.Fields("OrdreSection") Then
 'remplis la table impression_soumission
 'ajoute seulement la section
 sOrdreSection = rstPiece.Fields("OrdreSection")
 
 If m_eLangage = ANGLAIS Then
 sSection = "NomSectionEN"
 Else
 sSection = "NomSectionFR"
 End If
 
 Call rstTemp.Open("SELECT " & sSection & " FROM GrbSoumProjSectionElec WHERE IDSection = " & rstPiece.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
 
 'Ajoute la section dans la liste de pièces
 Call rstImpListePiece.AddNew
 
 rstImpListePiece.Fields("NoLigne") = iCompteurPiece
1  rstImpListePiece.Fields("IDSoumission") = sNoSoumission
 
 If Not IsNull(rstTemp.Fields(sSection)) Then
 rstImpListePiece.Fields("NomSection") = rstTemp.Fields(sSection)
 Else
 rstImpListePiece.Fields("NomSection") = " "
 End If
 
 Call rstImpListePiece.Update
 
 iCompteurPiece = iCompteurPiece + 1
 
 Call rstTemp.Close
 
 sSousSection = rstPiece.Fields("SousSection")
 
 If sSousSection = S_PAS_SOUS_SECTION Then
 sSousSection = " "
 End If
 
 Call rstImpListePiece.AddNew
 
 rstImpListePiece.Fields("NoLigne") = iCompteurPiece
 rstImpListePiece.Fields("IDSoumission") = sNoSoumission
 rstImpListePiece.Fields("SousSection") = sSousSection
 
 Call rstImpListePiece.Update
 
 iCompteurPiece = iCompteurPiece + 1
 Else
 'ajoute une soussection dans impression_soum
 If sSousSection <> sSousSectionRS Then
 sSousSection = sSousSectionRS
 
 'ajoute une sous-section dans impression_piece
 Call rstImpListePiece.AddNew
 
 rstImpListePiece.Fields("NoLigne") = iCompteurPiece
 rstImpListePiece.Fields("IDSoumission") = sNoSoumission
 rstImpListePiece.Fields("SousSection") = sSousSection
 
 Call rstImpListePiece.Update
 
 iCompteurPiece = iCompteurPiece + 1
 End If
 End If
 
 'Ajoute la pièce à la liste de pièces
 Call rstImpListePiece.AddNew
 
 rstImpListePiece.Fields("NoLigne") = iCompteurPiece
 rstImpListePiece.Fields("IDSoumission") = sNoSoumission
 rstImpListePiece.Fields("NumItem") = rstPiece.Fields("NumItem")
 rstImpListePiece.Fields("Qté") = rstPiece.Fields("Qté")
 
 If m_eLangage = ANGLAIS Then
 rstImpListePiece.Fields("Description") = rstPiece.Fields("Desc_EN")
 Else
 rstImpListePiece.Fields("Description") = rstPiece.Fields("Desc_FR")
 End If
 
 rstImpListePiece.Fields("Manufact") = rstPiece.Fields("Manufact")
 
 Call rstImpListePiece.Update
 
4 iCompteurPiece = iCompteurPiece + 1
4 End If
 
 'Prochaine enreg
4 Call rstPiece.MoveNext
4 Loop
 
4 Call rstImpListePiece.Close
 
 ''''''''''''''''''''''''''''''''''''''''''''''''''
 ' rapport liste piece, met dans l'ordre de ligne '
 ''''''''''''''''''''''''''''''''''''''''''''''''''
 
4 Call rstImpListePiece.Open("SELECT * FROM Grbimpression_Listepiece WHERE IDSoumission = '" & sNoSoumission & "' ORDER BY noligne", g_connData, adOpenDynamic, adLockOptimistic)
 
4 Set DR_Liste_piece.DataSource = rstImpListePiece
 
4 Call TraduireImpressionListePiece

4 DR_Liste_piece.Sections("Section4").Controls("lblTitreNoFacture").Visible = True
4 DR_Liste_piece.Sections("Section4").Controls("lblNoFacture").Visible = True

4 DR_Liste_piece.Sections("Section4").Controls("lblNoFacture").Caption = sNoFacture
 
 'Affiche la date
4  DR_Liste_piece.Sections("Section3").Controls("lbldate").Caption = ConvertDate(Date)
 
4  DR_Liste_piece.Sections("Section4").Controls("lblprojet").Caption = sNoProjet
 
 'affiche l 'entête
4  DR_Liste_piece.Sections("Section4").Controls("lblsoumission").Caption = sNoSoumission
 
4  DR_Liste_piece.Sections("section4").Controls("lbldescription").Caption = rstProjSoum.Fields("Description")
 
4  Call rstTemp.Open("SELECT NomClient FROM GrbClient WHERE IDClient = " & rstProjSoum.Fields("IDClient"), g_connData, adOpenDynamic, adLockOptimistic)
 
4  DR_Liste_piece.Sections("section4").Controls("lblclient").Caption = rstTemp.Fields("NomClient")
 
4  Call rstTemp.Close
 
4  Call rstTemp.Open("SELECT NomContact FROM GrbContact WHERE IDContact = " & rstProjSoum.Fields("IDContact"), g_connData, adOpenDynamic, adLockOptimistic)
 
50 DR_Liste_piece.Sections("section4").Controls("lblcontact").Caption = rstTemp.Fields("nomcontact")
 
50 Call rstTemp.Close
 
 'Affiche le rapport liste des pieces
 DR_Liste_piece.Orientation = rptOrientPortrait

 Call DR_Liste_piece.Show(vbModal)
 
 Call rstImpListePiece.Close
 Set rstImpListePiece = Nothing
 
 Set rstTemp = Nothing
 
 Screen.MousePointer = vbDefault

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "ImprimerListePieces", Err, Err.number, Err.Description
End Sub

Private Sub TraduireImpressionListePiece()
 
 On Error GoTo Oups
 
 If m_eLangage = ANGLAIS Then
 DR_Liste_piece.Sections("Section4").Controls("lblTitreProjet").Caption = "Project:"
 DR_Liste_piece.Sections("Section4").Controls("lblTitreSoumission").Caption = "Quote:"
 
 DR_Liste_piece.Sections("Section2").Controls("lblTitreQuantite").Caption = "Qty"
 DR_Liste_piece.Sections("Section2").Controls("lblTitreNoItem").Caption = "Item No."
 DR_Liste_piece.Sections("Section2").Controls("lblTitreManufacturier").Caption = "Manufacturer"
 DR_Liste_piece.Sections("Section2").Controls("lblTitreID").Caption = "ID #"
 
 DR_Liste_piece.Sections("Section3").Controls("lblNoPage").Caption = "Page %p of %P"
 Else
 DR_Liste_piece.Sections("Section4").Controls("lblTitreProjet").Caption = "Projet:"
  DR_Liste_piece.Sections("Section4").Controls("lblTitreSoumission").Caption = "Soumission:"
 
  DR_Liste_piece.Sections("Section2").Controls("lblTitreQuantite").Caption = "Qté"
  DR_Liste_piece.Sections("Section2").Controls("lblTitreNoItem").Caption = "No. Item"
  DR_Liste_piece.Sections("Section2").Controls("lblTitreManufacturier").Caption = "Manufacturier"
  DR_Liste_piece.Sections("Section2").Controls("lblTitreID").Caption = "# ID"
 
  DR_Liste_piece.Sections("Section3").Controls("lblNoPage").Caption = "Page %p de %P"
  End If

  Exit Sub

Oups:

10 wOups "frmProjSoumElec", "TraduireImpressionListePiece", Err, Err.number, Err.Description
End Sub

Private Sub TraduireImpressionSoumission()

 On Error GoTo Oups

 If m_eLangage = ANGLAIS Then
 If m_eType = TYPE_PROJET Then
 DR_SoumissionElec.Caption = "Electrical Project"
 DR_SoumissionElec.Sections("Section2").Controls("lblGrosTitre").Caption = "Electrical Project"
 Else
 DR_SoumissionElec.Caption = "Electrical Quote"
 DR_SoumissionElec.Sections("Section2").Controls("lblGrosTitre").Caption = "Electrical Quote"
 End If
 
 DR_SoumissionElec.Sections("Section2").Controls("lblTitreProjet").Caption = "Project :"
 DR_SoumissionElec.Sections("Section2").Controls("lblTitreSoumission").Caption = "Quote :"
  DR_SoumissionElec.Sections("Section2").Controls("lblTitreClient").Caption = "Client :"
  DR_SoumissionElec.Sections("Section2").Controls("lblTitreContact").Caption = "Contact :"

 '************************************************************************************************
 'SECTION MODIFIÉ PAR GAÉTAN GINGRAS LE 0  FÉVRIER 2010
 '************************************************************************************************
 
  DR_SoumissionElec.Sections("Section2").Controls("lblTitreQuantite").Caption = "Qty"
  DR_SoumissionElec.Sections("Section2").Controls("lblTitreNoItem").Caption = "Item No."
  DR_SoumissionElec.Sections("Section2").Controls("lblTitreDescription").Caption = "Description"
  DR_SoumissionElec.Sections("Section2").Controls("lblTitreManufacturier").Caption = "Manufacturer"
  'DR_SoumissionElec.Sections("Section2").Controls("lblTitrePrixListe").Caption = "Listed Price"
  'DR_SoumissionElec.Sections("Section2").Controls("lblTitreEscompte").Caption = "Discount"
DR_SoumissionElec.Sections("Section2").Controls("lblTitreCoutant").Caption = "Cost"
1 DR_SoumissionElec.Sections("Section2").Controls("lblTitreFournisseur").Caption = "Supplier"
 'DR_SoumissionElec.Sections("Section2").Controls("lblTitreTempsMontage").Caption = "Time"
 'DR_SoumissionElec.Sections("Section2").Controls("lblTitreMontage").Caption = "Fixing Time"
 DR_SoumissionElec.Sections("Section2").Controls("lblTitreTotal").Caption = "Total"

 'AJOUT PAR GAÉTAN GINGRAS 0  FÉVRIER 2010
 '****************************************************************************************
 DR_SoumissionElec.Sections("Section2").Controls("lbl_DateCommande").Caption = "Order Date"
 DR_SoumissionElec.Sections("Section2").Controls("lbl_DateReception").Caption = "Reception Date"
 '****************************************************************************************
 
 DR_SoumissionElec.Sections("Section5").Controls("lblTitreDessin").Caption = "Drafting :"
 DR_SoumissionElec.Sections("Section5").Controls("lblTitreFabrication").Caption = "Manufacturing :"
 DR_SoumissionElec.Sections("Section5").Controls("lblTitreAssemblage").Caption = "Assembling :"
 DR_SoumissionElec.Sections("Section5").Controls("lblTitreProgInterface").Caption = "Interface programming :"
 DR_SoumissionElec.Sections("Section5").Controls("lblTitreProgAutomate").Caption = "PLC programming :"
 DR_SoumissionElec.Sections("Section5").Controls("lblTitreProgRobot").Caption = "Robot programming :"
 DR_SoumissionElec.Sections("Section5").Controls("lblTitreVision").Caption = "Vision :"
DR_SoumissionElec.Sections("Section5").Controls("lblTitreTest").Caption = "Test :"
 DR_SoumissionElec.Sections("Section5").Controls("lblTitreInstallation").Caption = "Installation :"
 DR_SoumissionElec.Sections("Section5").Controls("lblTitreMiseService").Caption = "Activation :"
 DR_SoumissionElec.Sections("Section5").Controls("lblTitreFormation").Caption = "Training :"
 DR_SoumissionElec.Sections("Section5").Controls("lblTitreGestion").Caption = "Project management :"
 DR_SoumissionElec.Sections("Section5").Controls("lblTitreShipping").Caption = "Shipping :"

 DR_SoumissionElec.Sections("Section5").Controls("lblTitreTauxHoraire").Caption = "Rate / Hours"
1  DR_SoumissionElec.Sections("Section5").Controls("lblTitreTemps").Caption = "Time (Hour)"
 DR_SoumissionElec.Sections("Section5").Controls("lblTitreTotalPiece").Caption = "Parts Total:"
 DR_SoumissionElec.Sections("Section5").Controls("lblTitreImprevue").Caption = "Unforeseen:"
 DR_SoumissionElec.Sections("Section5").Controls("lblTitreTotalTemps").Caption = "Time Total:"
 DR_SoumissionElec.Sections("Section5").Controls("lblTitreGrandTotal").Caption = "Final Price:"
 
 DR_SoumissionElec.Sections("Section3").Controls("lblNoPage").Caption = "Page %p of %P"
 DR_SoumissionElec.Sections("Section5").Controls("lblTitrePrixReception").Caption = "Receiving up to date"
 DR_SoumissionElec.Sections("Section5").Controls("lblTitrePrixSoumission").Caption = "Quote Price"
 DR_SoumissionElec.Sections("Section5").Controls("lblTitreForfait").Caption = "Package Deal"
Else
 If m_eType = TYPE_PROJET Then
 DR_SoumissionElec.Caption = "Projet Électrique"
 DR_SoumissionElec.Sections("Section2").Controls("lblGrosTitre").Caption = "Projet Électrique"
Else
 DR_SoumissionElec.Caption = "Soumission Électrique"
 DR_SoumissionElec.Sections("Section2").Controls("lblGrosTitre").Caption = "Soumission Électrique"
 End If
 
DR_SoumissionElec.Sections("Section2").Controls("lblTitreProjet").Caption = "Projet:"
 DR_SoumissionElec.Sections("Section2").Controls("lblTitreSoumission").Caption = "Soumission:"
DR_SoumissionElec.Sections("Section2").Controls("lblTitreClient").Caption = "Client:"
 DR_SoumissionElec.Sections("Section2").Controls("lblTitreContact").Caption = "Contact:"
 
DR_SoumissionElec.Sections("Section2").Controls("lblTitreQuantite").Caption = "Qté"
3 DR_SoumissionElec.Sections("Section2").Controls("lblTitreNoItem").Caption = "No. Item"
 DR_SoumissionElec.Sections("Section2").Controls("lblTitreDescription").Caption = "Description"
 DR_SoumissionElec.Sections("Section2").Controls("lblTitreManufacturier").Caption = "Manufacturier"
 'DR_SoumissionElec.Sections("Section2").Controls("lblTitrePrixListe").Caption = "Prix listé"
 'DR_SoumissionElec.Sections("Section2").Controls("lblTitreEscompte").Caption = "Escompte"
 DR_SoumissionElec.Sections("Section2").Controls("lblTitreCoutant").Caption = "Coûtant"
 DR_SoumissionElec.Sections("Section2").Controls("lblTitreFournisseur").Caption = "Fournisseur"
 'DR_SoumissionElec.Sections("Section2").Controls("lblTitreTempsMontage").Caption = "Temps"
 'DR_SoumissionElec.Sections("Section2").Controls("lblTitreMontage").Caption = "Montage"
 DR_SoumissionElec.Sections("Section2").Controls("lblTitreTotal").Caption = "Total"

 'AJOUT PAR GAÉTAN GINGRAS 0  FÉVRIER 2010
 '****************************************************************************************
 DR_SoumissionElec.Sections("Section2").Controls("lbl_DateCommande").Caption = "Date commandé"
 DR_SoumissionElec.Sections("Section2").Controls("lbl_DateReception").Caption = "Date reçu"
 '****************************************************************************************

 '************************************************************************************************
 'FIN DE LA SECTION DE MODIFICATION
 '************************************************************************************************
 
 DR_SoumissionElec.Sections("Section5").Controls("lblTitreDessin").Caption = "Dessin :"
DR_SoumissionElec.Sections("Section5").Controls("lblTitreFabrication").Caption = "Fabrication :"
 DR_SoumissionElec.Sections("Section5").Controls("lblTitreAssemblage").Caption = "Assemblage :"
DR_SoumissionElec.Sections("Section5").Controls("lblTitreProgInterface").Caption = "Programmation d'interface :"
 DR_SoumissionElec.Sections("Section5").Controls("lblTitreProgAutomate").Caption = "Programmation d'automate :"
DR_SoumissionElec.Sections("Section5").Controls("lblTitreProgRobot").Caption = "Programmation de robot :"
 DR_SoumissionElec.Sections("Section5").Controls("lblTitreVision").Caption = "Vision :"
 DR_SoumissionElec.Sections("Section5").Controls("lblTitreTest").Caption = "Test :"
 DR_SoumissionElec.Sections("Section5").Controls("lblTitreInstallation").Caption = "Installation :"
DR_SoumissionElec.Sections("Section5").Controls("lblTitreMiseService").Caption = "Mise en service :"
4 DR_SoumissionElec.Sections("Section5").Controls("lblTitreFormation").Caption = "Formation du personnel :"
4 DR_SoumissionElec.Sections("Section5").Controls("lblTitreGestion").Caption = "Gestion du projet :"
4 DR_SoumissionElec.Sections("Section5").Controls("lblTitreShipping").Caption = "Expédition :"

4 DR_SoumissionElec.Sections("Section5").Controls("lblTitreTauxHoraire").Caption = "Taux Horaire"
4 DR_SoumissionElec.Sections("Section5").Controls("lblTitreTemps").Caption = "Temps (Heure)"
4 DR_SoumissionElec.Sections("Section5").Controls("lblTitreTotalPiece").Caption = "Total pièce:"
4 DR_SoumissionElec.Sections("Section5").Controls("lblTitreImprevue").Caption = "Imprévue:"
4 DR_SoumissionElec.Sections("Section5").Controls("lblTitreTotalTemps").Caption = "Total temps:"
4 DR_SoumissionElec.Sections("Section5").Controls("lblTitreGrandTotal").Caption = "Grand total:"
 
4 DR_SoumissionElec.Sections("Section3").Controls("lblNoPage").Caption = "Page %p de %P"
4 DR_SoumissionElec.Sections("Section5").Controls("lblTitrePrixReception").Caption = "Réception jusqu'à maintenant"
4  DR_SoumissionElec.Sections("Section5").Controls("lblTitrePrixSoumission").Caption = "$ Soumission"
4  DR_SoumissionElec.Sections("Section5").Controls("lblTitreForfait").Caption = "Forfait"
4  End If

4  Exit Sub

Oups:

4  wOups "frmProjSoumElec", "TraduireImpressionSoumission", Err, Err.number, Err.Description
End Sub

Private Sub cmdModifier_Click()

 On Error GoTo Oups

 Dim rstProjSoum As ADODB.Recordset
 Dim sUser As String

 Set m_collQteSupp = New Collection
 Set m_collDateSupp = New Collection
 Set m_collHeureSupp = New Collection
 Set m_collNoItemSupp = New Collection

 'Modifier une soumission
 If cmbProjSoum.ListIndex > -1 Then
 If Right$(txtNoProjSoum.Text, 2) = "99" Then
 If m_eType = TYPE_PROJET Then
 Call MsgBox("Ce projet ne peut pas être modifié!", vbOKOnly, "Erreur")
  Else
  Call MsgBox("Cette soumission ne peut pas être modifiée!", vbOKOnly, "Erreur")
  End If
 
  Exit Sub
  End If

  If m_eType = TYPE_SOUMISSION Then
  If VerifierSiDejaProjet = True Then
  Call MsgBox("Vous ne pouvez pas modifier cette soumission, le projet a déjà été créé!", vbOKOnly, "Erreur")

 Exit Sub
End If
 End If

 Set rstProjSoum = New ADODB.Recordset

 Call rstProjSoum.Open("SELECT * FROM GrbProjSoum WHERE IDProjSoum ='" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 If rstProjSoum.Fields("Ouvert") = False Or rstProjSoum.Fields("Verrouillé") = True Then
 If rstProjSoum.Fields("Ouvert") = False Then
 If m_eType = TYPE_PROJET Then
 Call MsgBox("Ce projet est fermé!", vbOKOnly, "Erreur")
 Else
 Call MsgBox("Cette soumission est fermée!", vbOKOnly, "Erreur")
 End If
 Else
 If m_eType = TYPE_PROJET Then
 Call MsgBox("Ce projet est verrouillé!", vbOKOnly, "Erreur")
 Else
 Call MsgBox("Cette soumission est verrouillée!", vbOKOnly, "Erreur")
 End If
 End If
 
1  Call rstProjSoum.Close
 Set rstProjSoum = Nothing
 
 Exit Sub
 End If
 
 Call rstProjSoum.Close
 Set rstProjSoum = Nothing
 
 If VerifierSiOuvert(sUser) = False Then
 Screen.MousePointer = vbHourglass
 
 'Débarre les champs
 Call BarrerChamps(False)
 
 'Pour pouvoir afficher le dernier enregistrement affiché quand la personne va
 'enregistrer ou annuler
 m_sAncienProjSoum = txtNoProjSoum.Text
 
 m_bModeAjout = False
 m_bModeAffichage = False
 
 'Rapetisse le listview de la soumission pour afficher le lvwPiece
 lvwSoumission.Height = lvwSoumission.Height * 0.49
 lvwSoumission.Top = lvwPieces.Top + lvwPieces.Height + (lvwSoumission.Height * 0.02)
 
 Call RemplirProjSoum
 
 Call AfficherControles(MODE_AJOUT_MODIF)
 
 Call UpdateOrdre
 
 'On recalcul le prix total
 Call CalculerPrix
 
 Call lvwSoumission.Refresh
 
 Call OuvrirProjSoum(True)
 
 Screen.MousePointer = vbDefault
Else
If m_eType = TYPE_PROJET Then
 Call MsgBox("Ce projet est en modification par " & sUser & "!", vbOKOnly, "Erreur")
 Else
 Call MsgBox("Cette soumission est en modification par " & sUser & "!", vbOKOnly, "Erreur")
 End If
 End If
End If

Exit Sub

Oups:

wOups "frmProjSoumElec", "cmdModifier_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdsupprimer_Click()

 On Error GoTo Oups

 Dim iReponse As Integer
 Dim rstProjSoum As ADODB.Recordset
 Dim rstProjet As ADODB.Recordset
 Dim sSoumission As String
 Dim sUser As String
 Dim iExtension As Integer

 'Si il y a des enregistrements
 If cmbProjSoum.ListCount > 0 Then
 If Right$(txtNoProjSoum.Text, 2) = "99" Then
 If m_eType = TYPE_PROJET Then
 Call MsgBox("Vous ne pouvez pas supprimer ce projet!", vbOKOnly, "Erreur")
  Else
  Call MsgBox("Vous ne pouvez pas supprimer cette soumission!", vbOKOnly, "Erreur")
  End If

  Exit Sub
  End If

  Set rstProjSoum = New ADODB.Recordset

  Call rstProjSoum.Open("SELECT * FROM GrbProjSoum WHERE IDProjSoum = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

  If rstProjSoum.Fields("Ouvert") = False Or rstProjSoum.Fields("Verrouillé") = True Then
 If rstProjSoum.Fields("Ouvert") = False Then
 If m_eType = TYPE_PROJET Then
 Call MsgBox("Ce projet est fermé!", vbOKOnly, "Erreur")
 Else
 Call MsgBox("Cette soumission est fermée!", vbOKOnly, "Erreur")
 End If
 Else
 If m_eType = TYPE_PROJET Then
 Call MsgBox("Ce projet est verrouillé!", vbOKOnly, "Erreur")
 Else
 Call MsgBox("Cette soumission est verrouillée!", vbOKOnly, "Erreur")
 End If
 End If

 Call rstProjSoum.Close
 Set rstProjSoum = Nothing

 Exit Sub
 End If

 Call rstProjSoum.Close
 Set rstProjSoum = Nothing
 
1  If m_eType = TYPE_SOUMISSION Then
 If VerifierSiDejaProjet = True Then
 Call MsgBox("Vous ne pouvez pas supprimer cette soumission, le projet a déjà été créé!", vbOKOnly, "Erreur")
 
 Exit Sub
 End If
 End If
 
 If VerifierSiOuvert(sUser) = False Then
 'Valider le choix
 If m_eType = TYPE_PROJET Then
 iReponse = MsgBox("Voulez-vous vraiment EFFACER LE PROJET " & txtNoProjSoum.Text & "?", vbYesNo)

 If iReponse = vbYes Then
 Call frmValiderSuppression.Afficher(True, txtNoProjSoum.Text, Me)

 If m_bValide = True Then
 iReponse = vbYes
 Else
 iReponse = vbNo
 End If
 End If
 Else
 iReponse = MsgBox("Voulez-vous vraiment EFFACER LA SOUMISSION " & txtNoProjSoum.Text & "?", vbYesNo)

 If iReponse = vbYes Then
 Call frmValiderSuppression.Afficher(False, txtNoProjSoum.Text, Me)

 If m_bValide = True Then
 iReponse = vbYes
 Else
 iReponse = vbNo
 End If
 End If
 End If
 
 'S'il veut vraiment effacer
 If iReponse = vbYes Then
 'Si c'est un projet
 If m_eType = TYPE_PROJET Then
 Set rstProjet = New ADODB.Recordset

 Call rstProjet.Open("SELECT IDSoumission FROM GrbProjetElec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

 If Not IsNull(rstProjet.Fields("IDSoumission")) Then
 sSoumission = rstProjet.Fields("IDSoumission")
 Else
 sSoumission = vbNullString
 End If

 Call rstProjet.Close
 Set rstProjet = Nothing

 'Efface les pièces
 Call g_connData.Execute("DELETE * FROM GrbProjet_Pieces WHERE IDProjet = '" & txtNoProjSoum.Text & "' AND Type = 'E'")

 If IsNumeric(Right$(txtNoProjSoum.Text, 2)) Then
 iExtension = CInt(Right$(txtNoProjSoum.Text, 2))
4 Else
4 iExtension = 0
4 End If

4 If (iExtension >= 60 And iExtension <= 79) Or (iExtension >= 80 And iExtension <= 98) Then
4 Set rstProjSoum = New ADODB.Recordset

4 Call rstProjSoum.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

4 Call g_connData.Execute("DELETE * FROM GrbProjet_Pieces WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & rstProjSoum.Fields("LiaisonChargeable") & "' AND Provenance = '" & iExtension & "'")

4 Call CalculerTotalRecordset(Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & rstProjSoum.Fields("LiaisonChargeable"))
 
4 Call rstProjSoum.Close
4 Set rstProjSoum = Nothing
4 End If
 
 'Efface les modifications
4  Call g_connData.Execute("DELETE * FROM GrbProjet_Modif WHERE IDProjet = '" & txtNoProjSoum.Text & "' AND Type = 'E'")
 
 'Efface la soumission
4  Call g_connData.Execute("DELETE * FROM GrbProjetElec WHERE IDProjet = '" & txtNoProjSoum.Text & "'")

 'Efface la soumission dans la table GrbProjSoum
4  Call g_connData.Execute("DELETE * FROM GrbProjSoum WHERE IDProjSoum = '" & txtNoProjSoum.Text & "'")

4  Set rstProjSoum = New ADODB.Recordset

4  Call rstProjSoum.Open("SELECT Ouvert FROM GrbProjSoum WHERE IDProjSoum = '" & sSoumission & "'", g_connData, adOpenDynamic, adLockOptimistic)

4  If Not rstProjSoum.EOF Then
4  rstProjSoum.Fields("Ouvert") = True

4  Call rstProjSoum.Update
50 End If

 Call rstProjSoum.Close
 Set rstProjSoum = Nothing
 Else
 'Efface les pièces
 Call g_connData.Execute("DELETE * FROM GrbSoumission_Pieces WHERE IDSoumission = '" & txtNoProjSoum.Text & "' AND Type = 'E'")
 
 'Efface les modifications
 Call g_connData.Execute("DELETE * FROM GrbSoumission_Modif WHERE IDSoumission = '" & txtNoProjSoum.Text & "' AND Type = 'E'")
 
 'Efface la soumission
 Call g_connData.Execute("DELETE * FROM GrbSoumissionElec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'")
 
 'Efface la soumission dans la table GrbProjSoum
 Call g_connData.Execute("DELETE * FROM GrbProjSoum WHERE IDProjSoum = '" & txtNoProjSoum.Text & "'")
 End If

 If m_eType = TYPE_PROJET Then
 Call RecreerProjetCumulatif
 Else
5  Call RecreerSoumissionCumulatif
5  End If
 
 'Affiche la premiere soumission
5  Call AfficherProjSoum(vbNullString)
5  End If
5  Else
5  If m_eType = TYPE_PROJET Then
5  Call MsgBox("Ce projet est en modification par " & sUser & "!", vbOKOnly, "Erreur")
5  Else
60 Call MsgBox("Cette soumission est en modification par " & sUser & "!", vbOKOnly, "Erreur")
  End If
  End If
  End If

  Exit Sub

Oups:

  wOups "frmProjSoumElec", "cmdSupprimer_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups

 Call Unload(frmChoixProjSoum)
 
 m_eLangage = FRANCAIS
 
 cmdAnglaisFrancais.Caption = "Anglais"

 'Initialise le tri à PIECE_GRB
 cmbTri.ListIndex = I_CMB_PIECE
 
 'Donne accès aux boutons selon le groupe
 Call ActiverBoutonsGroupe

 'Initialisation au mode inactif
 m_eMode = MODE_INACTIF
 
 'Rempli le combo des clients
 Call RemplirComboClients(vbNullString)
 
 'Rempli le combo des contacts
 Call RemplirComboSections
 
 'Rempli le combo des catégories de pièce
 Call RemplirComboCategoriesPieces
 
 cmbOuvertFerme.ListIndex = I_CMB_OUVERT
 
  If m_eType = TYPE_PROJET Then
  cmbChoix.ListIndex = I_IDX_PROJET
  Else
  cmbChoix.ListIndex = I_IDX_SOUMISSION
  End If

  Exit Sub

Oups:

  wOups "frmProjSoumElec", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub RemplirColonnes()

 On Error GoTo Oups
 
 'Méthode pour afficher les colonnes selon le groupe de sécurité.
 
 'Ceux qui n'ont pas le droit de modifier les soumissions ou les projets, n'ont pas
 'le droit de voir les prix, donc il faut cacher les colonnes
 Dim bModif As Boolean
 Dim bCacherPrix As Boolean
 
 'Si le type d'affichage est "Projet"
 If m_eType = TYPE_PROJET Then
 'Si l'utilisateur n'a pas le droit de modification sur les projets
 If m_bModifProj = False Then
 'On cache les prix
 bCacherPrix = True
 End If
 Else
 'Si l'utilisateur n'a pas le droit de modification sur les soumissions
 If m_bModifSoum = False Then
 'On cache les prix
 bCacherPrix = True
 End If
  End If
 
 'Si on cache les prix
  If bCacherPrix = True Then
  m_bDroitPrix = False
 
 'Si les bonnes colonnes sont déjà toutes affichées
  If lvwSoumission.ColumnHeaders.count = 12 Then
  Exit Sub
  End If
  Else
  m_bDroitPrix = True

 'Si les colonnes sont déjà toute là
If lvwSoumission.ColumnHeaders.count = 1 Then
Exit Sub
 End If
End If
 
 'Il faut enlever les colonnes avant d'en ajouter d'autres
Call lvwSoumission.ColumnHeaders.Clear
 
Call lvwSoumission.ColumnHeaders.Add(, , "Qté", 650.1418)
Call lvwSoumission.ColumnHeaders.Add(, , "No. Item", 1830.0474)
Call lvwSoumission.ColumnHeaders.Add(, , "Description", 3809.746)
Call lvwSoumission.ColumnHeaders.Add(, , "Manufacturier", 1154.8348)
 
If bCacherPrix = False Then
 Call lvwSoumission.ColumnHeaders.Add(, , "Prix listé", 920.1261, vbRightJustify)
 Call lvwSoumission.ColumnHeaders.Add(, , "Escompte", 884.9765, vbRightJustify)
Call lvwSoumission.ColumnHeaders.Add(, , "Prix net", 920.1261, vbRightJustify)
End If
 
 Call lvwSoumission.ColumnHeaders.Add(, , "Distributeur", 1005.1655)
Call lvwSoumission.ColumnHeaders.Add(, , "Temps", 824.882)
 Call lvwSoumission.ColumnHeaders.Add(, , "Montage", 824.882)
 
If bCacherPrix = False Then
 Call lvwSoumission.ColumnHeaders.Add(, , "TOTAL", 1099.8426, vbRightJustify)
1  Call lvwSoumission.ColumnHeaders.Add(, , "Profit", 920.1261, vbRightJustify)
 End If

 Call lvwSoumission.ColumnHeaders.Add(, , "Commentaire", 1000)

If m_eType = TYPE_PROJET Then
 Call lvwSoumission.ColumnHeaders.Add(, , "ID", 1440)

 If bCacherPrix = False Then
 Call lvwSoumission.ColumnHeaders.Add(, , "Facturation", 1440)
 End If

 Call lvwSoumission.ColumnHeaders.Add(, , "Date Commande", 1440)
 Call lvwSoumission.ColumnHeaders.Add(, , "Date Requise", 1440)
 Call lvwSoumission.ColumnHeaders.Add(, , "Commandé par", 1440)
 Call lvwSoumission.ColumnHeaders.Add(, , "No Séquentiel", 1440)
End If

2  Call lvwSoumission.ColumnHeaders.Add(, , "Provenance", 1440)

Exit Sub

Oups:

2  wOups "frmProjSoumElec", "RemplirColonnes", Err, Err.number, Err.Description
End Sub

Private Sub BarrerChamps(ByVal bBarrer As Boolean)

 On Error GoTo Oups

 'Méthode qui barre ou débarre les champs d'après la variable bBarrer
 txtProjet.Locked = bBarrer
 txtNbreManuel.Locked = bBarrer
 txtPrixManuel.Locked = bBarrer
 picApprob.Enabled = Not bBarrer
 txtCheminPhotos.Locked = bBarrer

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "BarrerChamps", Err, Err.number, Err.Description
End Sub

Private Sub ViderChamps()

 On Error GoTo Oups

 'Méthode qui initialise les champs
 txtClient.Text = vbNullString
 txtcontact.Text = vbNullString
 txtProjet.Text = vbNullString
 txtNbreManuel.Text = 0
 txtPrixManuel.Text = 0
 txtTransport.Text = vbNullString
 txtPrixReception.Text = Conversion("0", MODE_ARGENT)
 txtPrixSoumission.Text = Conversion("0", MODE_ARGENT)
 chkCSA.Value = vbUnchecked
 chkCUL.Value = vbUnchecked
  chkUL.Value = vbUnchecked
  chkCUR.Value = vbUnchecked
  chkUR.Value = vbUnchecked
  chkCE.Value = vbUnchecked
  txtPrixTotal.Text = 0
  txtProfit.Text = 0
  txtDelais.Text = vbNullString
  txtCommission.Text = 0
10 txtNoSoumission.Text = vbNullString
txtCheminPhotos.Text = vbNullString
txtForfait.Text = vbNullString
lblForfaitInitiale.Caption = vbNullString
 
cmbtransport.ListIndex = I_TRANS_FAB_GRANBY

cmbclient.ListIndex = -1

m_bSansTemps = False
lblPasTemps.Visible = False
tmrTemps.Enabled = False
 
Call lvwSoumission.ListItems.Clear

Exit Sub

Oups:

wOups "frmProjSoumElec", "ViderChamps", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboProjSoum(ByVal sNoProjSoum As String)

 On Error GoTo Oups

 'Rempli le combo des soumissions
 Dim rstProjSoum As ADODB.Recordset
 
 'Il faut vider le combo avant de le remplir
 Call cmbProjSoum.Clear
 
 Set rstProjSoum = New ADODB.Recordset
 
 'Ouvre le recordset selon le type
 If m_eType = TYPE_PROJET Then
 If cmbOuvertFerme.ListIndex = I_CMB_OUVERT Then
 Call rstProjSoum.Open("SELECT IDProjet FROM GrbProjetElec INNER JOIN GrbProjSoum ON GrbProjetElec.IDProjet = GrbProjSoum.IDProjSoum WHERE Ouvert = True ORDER BY IDProjet DESC", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstProjSoum.Open("SELECT IDProjet FROM GrbProjetElec ORDER BY IDProjet DESC", g_connData, adOpenDynamic, adLockOptimistic)
 End If
 Else
  If cmbOuvertFerme.ListIndex = I_CMB_OUVERT Then
  Call rstProjSoum.Open("SELECT IDSoumission FROM GrbSoumissionElec INNER JOIN GrbProjSoum ON GrbSoumissionElec.IDSoumission = GrbProjSoum.IDProjSoum WHERE Ouvert = True ORDER BY IDSoumission DESC", g_connData, adOpenDynamic, adLockOptimistic)
  Else
  Call rstProjSoum.Open("SELECT IDSoumission FROM GrbSoumissionElec ORDER BY IDSoumission DESC", g_connData, adOpenDynamic, adLockOptimistic)
  End If
  End If
 
 'Tant que ce n'est pas la fin des enregistrements
  Do While Not rstProjSoum.EOF
 'On met le numéro de la soumission dans le combo des soumissions
  If m_eType = TYPE_PROJET Then
 Call cmbProjSoum.AddItem(rstProjSoum.Fields("IDProjet"))
1 Else
 Call cmbProjSoum.AddItem(rstProjSoum.Fields("IDSoumission"))
 End If
 
 Call rstProjSoum.MoveNext
Loop
 
Call rstProjSoum.Close
Set rstProjSoum = Nothing
 
 'Si le combo n'est pas vide, on sélectionne l'item voulu ou le premier
If cmbProjSoum.ListCount > 0 Then
 'Si il y a un numéro de projet
 If sNoProjSoum <> vbNullString Then
 'On le sélectionne dans le combo
 Call RechercherProjSoum(sNoProjSoum)
 Else
 'Sinon, on sélectionne le premier
 cmbProjSoum.ListIndex = 0
 End If
 End If

Exit Sub

Oups:

 wOups "frmProjSoumElec", "RemplirComboProjSoum", Err, Err.number, Err.Description
End Sub

Private Sub CalculerPrixReel(ByVal sNoItem As String)

 On Error GoTo Oups

 Dim rstPieceFRS As ADODB.Recordset
 Dim rstConfig As ADODB.Recordset
 Dim sPrixCalcul As String
 Dim sTauxUSA As String
 Dim sTauxSPA As String
 
 Set rstConfig = New ADODB.Recordset

 Call rstConfig.Open("SELECT TauxAmericain, TauxEspagnol FROM GrbConfig", g_connData, adOpenForwardOnly, adLockReadOnly)

 sTauxUSA = rstConfig.Fields("TauxAmericain")
 sTauxSPA = rstConfig.Fields("TauxEspagnol")

 Call rstConfig.Close
  Set rstConfig = Nothing
 
  Set rstPieceFRS = New ADODB.Recordset
 
  rstPieceFRS.CursorLocation = adUseServer
 
  Call rstPieceFRS.Open("SELECT PrixReel, PRIX_NET, PRIX_SP, DeviseMonétaire FROM GrbPiecesFRS WHERE PIECE = '" & Replace(sNoItem, "'", "''") & "' AND Type = 'E'", g_connData, adOpenDynamic, adLockOptimistic)
 
  Do While Not rstPieceFRS.EOF
  If rstPieceFRS.Fields("PRIX_NET") <> vbNullString Then
  sPrixCalcul = rstPieceFRS.Fields("PRIX_NET")
  Else
 If rstPieceFRS.Fields("PRIX_SP") <> vbNullString Then
 sPrixCalcul = rstPieceFRS.Fields("PRIX_SP")
 End If
 End If

 sPrixCalcul = Replace(sPrixCalcul, ".", ",")

 If rstPieceFRS.Fields("DeviseMonétaire") = "USA" Then
 rstPieceFRS.Fields("PrixReel") = Conversion(CStr(Round(CDbl(sPrixCalcul) / CDbl(sTauxUSA), 4)), MODE_DECIMAL, 4)
 Else
 If rstPieceFRS.Fields("DeviseMonétaire") = "SPA" Then
 rstPieceFRS.Fields("PrixReel") = Conversion(CStr(Round(CDbl(sPrixCalcul) / CDbl(sTauxSPA), 4)), MODE_DECIMAL, 4)
 Else
 rstPieceFRS.Fields("PrixReel") = sPrixCalcul
 End If
 End If

 Call rstPieceFRS.Update

 Call rstPieceFRS.MoveNext
 Loop

Call rstPieceFRS.Close
 Set rstPieceFRS = Nothing

1  Exit Sub

Oups:

 wOups "frmProjSoumElec", "CalculerPrixReel", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListViewFournisseur()

 On Error GoTo Oups

 'Rempli le listview des distributeurs pour une pièce choisie
 Dim rstPieceFRS As ADODB.Recordset
 Dim rstContact As ADODB.Recordset
 Dim rstFRS As ADODB.Recordset
 Dim rstInv As ADODB.Recordset
 Dim itmFRS As ListItem
 Dim iCompteur As Integer
 Dim iNoClient As Integer
 Dim bAjouterDP As Boolean
 Dim sDevise As String
 Dim lColor As Long
 
  Set rstPieceFRS = New ADODB.Recordset
  Set rstContact = New ADODB.Recordset
  Set rstFRS = New ADODB.Recordset
 
 'vide le lister
  Call lvwfournisseur.ListItems.Clear

  If m_bPieceInutile = True Or m_bChangementFRS = True Then
  Call CalculerPrixReel(Trim$(lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE)))
  Else
  If m_bRecherchePiece = True Then
 Call CalculerPrixReel(Trim$(lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM)))
1 Else
 Call CalculerPrixReel(Trim$(lvwPieces.SelectedItem.SubItems(I_COL_PIECES_NO_ITEM)))
 End If
End If
 
Call rstFRS.Open("SELECT IDFRS FROM GrbFournisseur WHERE NomFournisseur = 'FOURNI PAR LE CLIENT'", g_connData, adOpenDynamic, adLockOptimistic)
 
iNoClient = rstFRS.Fields("IDFRS")

Call rstFRS.Close
Set rstFRS = Nothing
 
If m_bPieceInutile = True Or m_bChangementFRS = True Then
 Call rstPieceFRS.Open("SELECT GrbPiecesFRS.*, GrbFournisseur.NomFournisseur FROM GrbPiecesFRS INNER JOIN GrbFournisseur ON GrbPiecesFRS.IDFRS = GrbFournisseur.IDFRS WHERE PIECE = '" & Trim$(Replace(lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE), "'", "''")) & "' AND Type = 'E' ORDER BY CDbl(PrixReel)", g_connData, adOpenDynamic, adLockOptimistic)
Else
If m_bRecherchePiece = True Then
 Call rstPieceFRS.Open("SELECT GrbPiecesFRS.*, GrbFournisseur.NomFournisseur FROM GrbPiecesFRS INNER JOIN GrbFournisseur ON GrbPiecesFRS.IDFRS = GrbFournisseur.IDFRS WHERE PIECE = '" & Trim$(Replace(lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM), "'", "''")) & "' AND Type = 'E' ORDER BY CDbl(PrixReel)", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstPieceFRS.Open("SELECT GrbPiecesFRS.*, GrbFournisseur.NomFournisseur FROM GrbPiecesFRS INNER JOIN GrbFournisseur ON GrbPiecesFRS.IDFRS = GrbFournisseur.IDFRS WHERE PIECE = '" & Trim$(Replace(lvwPieces.SelectedItem.SubItems(I_COL_PIECES_NO_ITEM), "'", "''")) & "' AND Type = 'E' ORDER BY CDbl(PrixReel)", g_connData, adOpenDynamic, adLockOptimistic)
 End If
End If

 'Tant qu'il y a des fournisseur de la pièce, on ajoute dans le ListView
 Do While Not rstPieceFRS.EOF
1  If m_bPieceInutile = True Then
 If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_BRUN Then
 If rstPieceFRS.Fields("IDFRS") = iNoClient Then
 Call rstPieceFRS.MoveNext

 If rstPieceFRS.EOF Then
 Exit Do
 End If
 End If
 End If
 End If

 'on change la couleur de l'enregistrement selon la devise monétaire.
 'CAN = noir, USA ou ESP = bleu
 If rstPieceFRS.Fields("DeviseMonétaire") = "CAN" Then
 sDevise = "CAN"
 lColor = COLOR_NOIR
Else
 If rstPieceFRS.Fields("DeviseMonétaire") = "USA" Then
 sDevise = "USA"
 lColor = COLOR_BLEU
 Else
 sDevise = "SPA"
 lColor = COLOR_BLEU
 End If
End If
 
3 Set itmFRS = lvwfournisseur.ListItems.Add
 
 itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = " "
 itmFRS.SubItems(I_COL_FRS_ESCOMPTE) = " "
 itmFRS.SubItems(I_COL_FRS_PRIX_NET) = " "
 itmFRS.SubItems(I_COL_FRS_PRIX_SP) = " "
 
 'Nom du FRS
 itmFRS.Text = rstPieceFRS.Fields("NomFournisseur")
 
 itmFRS.Tag = rstPieceFRS.Fields("IDFRS")

 itmFRS.ForeColor = lColor
 
 'Personne ressource
 If Not IsNull(rstPieceFRS.Fields("PERS_RESS")) Then
 If Trim(rstPieceFRS.Fields("PERS_RESS")) <> vbNullString Then
 Call rstContact.Open("SELECT NomContact FROM GrbContact WHERE IDContact = " & rstPieceFRS.Fields("PERS_RESS"), g_connData, adOpenDynamic, adLockOptimistic)
 
 If Not rstContact.EOF Then
 itmFRS.SubItems(I_COL_FRS_PERS_RESS) = rstContact.Fields("NomContact")

 itmFRS.ListSubItems(I_COL_FRS_PERS_RESS).ForeColor = lColor
 End If

 Call rstContact.Close
 End If
 End If
  
 'Date
 If Not IsNull(rstPieceFRS.Fields("Date")) Then
 itmFRS.SubItems(I_COL_FRS_DATE) = rstPieceFRS.Fields("Date")
4 Else
4 itmFRS.SubItems(I_COL_FRS_DATE) = vbNullString
4 End If

4 itmFRS.ListSubItems(I_COL_FRS_DATE).ForeColor = lColor
 
 'Entrer par
4 If Not IsNull(rstPieceFRS.Fields("Entrer_Par")) Then
4 itmFRS.SubItems(I_COL_FRS_ENTRER_PAR) = rstPieceFRS.Fields("ENTRER_PAR")
4 Else
4 itmFRS.SubItems(I_COL_FRS_ENTRER_PAR) = vbNullString
4 End If

4 itmFRS.ListSubItems(I_COL_FRS_ENTRER_PAR).ForeColor = lColor
  
 'Valide
4 If Not IsNull(rstPieceFRS.Fields("Valide")) Then
4  itmFRS.SubItems(I_COL_FRS_VALIDE) = rstPieceFRS.Fields("Valide")
4  Else
4  itmFRS.SubItems(I_COL_FRS_VALIDE) = vbNullString
4  End If
 
4  itmFRS.ListSubItems(I_COL_FRS_VALIDE).ForeColor = lColor
 
 'Prix listé
4  If rstPieceFRS.Fields("PRIX_LIST") <> vbNullString Then
4  itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = Conversion(Round(Replace(rstPieceFRS.Fields("PRIX_LIST"), ".", ","), 4), MODE_ARGENT, 4)
4  End If

50 itmFRS.ListSubItems(I_COL_FRS_PRIX_LIST).ForeColor = lColor
 
 'Escompte
5 If rstPieceFRS.Fields("ESCOMPTE") <> vbNullString Then
 itmFRS.SubItems(I_COL_FRS_ESCOMPTE) = Conversion(Replace(Replace(rstPieceFRS.Fields("ESCOMPTE"), "_", vbNullString), ".", ",") * 100, MODE_POURCENT)
 End If

 itmFRS.ListSubItems(I_COL_FRS_ESCOMPTE).ForeColor = lColor

 'Prix net
 If rstPieceFRS.Fields("PRIX_NET") <> vbNullString Then
 itmFRS.SubItems(I_COL_FRS_PRIX_NET) = Conversion(Round(Replace(rstPieceFRS.Fields("PRIX_NET"), ".", ","), 4), MODE_ARGENT, 4)
 End If

 itmFRS.ListSubItems(I_COL_FRS_PRIX_NET).ForeColor = lColor
 
 'Prix spécial
 If rstPieceFRS.Fields("PRIX_SP") <> vbNullString Then
 itmFRS.SubItems(I_COL_FRS_PRIX_SP) = Conversion(Round(rstPieceFRS.Fields("PRIX_SP"), 4), MODE_ARGENT, 4)
 End If

5  itmFRS.ListSubItems(I_COL_FRS_PRIX_SP).ForeColor = lColor

 'Quoter
5  If rstPieceFRS.Fields("QUOTER") = True Then
5  itmFRS.SubItems(I_COL_FRS_QUOTER) = "Oui"
5  Else
5  itmFRS.SubItems(I_COL_FRS_QUOTER) = "Non"
5  End If

5  itmFRS.ListSubItems(I_COL_FRS_QUOTER).ForeColor = lColor

5  If rstPieceFRS.Fields("IDFRS") = 71 Then  'Si le fournisseur est "SOLUTION GRB Inc."
60 Set rstInv = New ADODB.Recordset

  Call rstInv.Open("SELECT * FROM GrbInventaireElec WHERE TRIM(NoItem) = '" & Trim(rstPieceFRS.Fields("PIECE")) & "'", g_connData, adOpenForwardOnly, adLockReadOnly)
 
  If Not rstInv.EOF Then
  If Not IsNull(rstInv.Fields("QuantitéStock")) Then
  itmFRS.SubItems(I_COL_FRS_STOCK) = rstInv.Fields("QuantitéStock")
  Else
  itmFRS.SubItems(I_COL_FRS_STOCK) = 0
  End If
  End If

  Call rstInv.Close
  Set rstInv = Nothing
  End If

 'Pour garder en mémoire le prix d'origine, je le mets dans le
 'tag de la colonne Prix Listé
6  If itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = vbNullString Then
6  itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = " "
6  End If
 
6  If rstPieceFRS.Fields("PRIX_NET") <> vbNullString Then
6  If rstPieceFRS.Fields("PRIX_LIST") = "0,00" Or rstPieceFRS.Fields("PRIX_LIST") = "0" Then
6  itmFRS.ListSubItems(I_COL_FRS_PRIX_LIST).Tag = Replace(rstPieceFRS.Fields("PRIX_NET"), ".", ",")
6  Else
6  itmFRS.ListSubItems(I_COL_FRS_PRIX_LIST).Tag = Replace(rstPieceFRS.Fields("PRIX_LIST"), ".", ",")
70 End If
  Else
  itmFRS.ListSubItems(I_COL_FRS_PRIX_LIST).Tag = Replace(rstPieceFRS.Fields("PRIX_SP"), ".", ",")
  End If

  If itmFRS.SubItems(I_COL_FRS_ENTRER_PAR) = vbNullString Then
  itmFRS.SubItems(I_COL_FRS_ENTRER_PAR) = " "
  End If

  itmFRS.ListSubItems(I_COL_FRS_ENTRER_PAR).Tag = rstPieceFRS.Fields("NoEnreg")

  If itmFRS.SubItems(I_COL_FRS_PERS_RESS) = "" Then
  itmFRS.SubItems(I_COL_FRS_PERS_RESS) = " "
  End If

  itmFRS.ListSubItems(I_COL_FRS_PERS_RESS).Tag = sDevise
 
   Call rstPieceFRS.MoveNext
   Loop
 
 'Ferme la table
7  Call rstPieceFRS.Close
7  Set rstPieceFRS = Nothing

7  Set rstContact = Nothing

7  If m_bPieceInutile = False Then
7  If lvwSoumission.ListItems.count > 0 Then
7  If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor <> COLOR_BRUN Then
80 bAjouterDP = True
  Else
  If m_bChangementFRS = False Then
  bAjouterDP = True
  End If
  End If
  Else
  bAjouterDP = True
  End If
  Else
  If m_bChangementFRS = True Then
  If lvwSoumission.ListItems.count > 0 Then
   If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor <> COLOR_BRUN Then
   bAjouterDP = True
   End If
   Else
8  bAjouterDP = True
8  End If
8  End If
8  End If

90 If bAjouterDP = True Then
  Set itmFRS = lvwfournisseur.ListItems.Add

  itmFRS.Text = "CHOISIR ULTÉRIEUREMENT"

  itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = " "
  itmFRS.SubItems(I_COL_FRS_ESCOMPTE) = " "
  itmFRS.SubItems(I_COL_FRS_PRIX_NET) = " "
  itmFRS.SubItems(I_COL_FRS_PRIX_SP) = " "
  End If

  Exit Sub

Oups:

  wOups "frmProjSoumElec", "RemplirListViewFournisseur", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListViewPieces()

 On Error GoTo Oups

 'Rempli le listview des pièces selon la catégorie de pièce choisit
 Dim rstPieces As ADODB.Recordset
 Dim itmPieces As ListItem
 Dim sCategorie As String
 Dim sTri As String
 Dim sOrderBy As String
 Dim bDebut As Boolean
 Dim iIndex As Integer
 
 sTri = m_sTri
 
 Select Case cmbTri.ListIndex
 Case I_CMB_PIECE_GRB: sOrderBy = "PIECE_GRB"
 Case I_CMB_PIECE: sOrderBy = "PIECE"
  Case I_CMB_FABRICANT: sOrderBy = "FABRICANT"
  Case I_CMB_DESCR_FR: sOrderBy = "DESC_FR"
  Case I_CMB_DESCR_EN: sOrderBy = "DESC_EN"
  End Select
 
 'Il faut vider le ListView avant de le remplir
  Call lvwPieces.ListItems.Clear
 
  sCategorie = Replace(cmbPieces.Text, "'", "''")
 
 'On ouvre un recordset selon la table choisie
  Set rstPieces = New ADODB.Recordset
 
  Call rstPieces.Open("SELECT * FROM GrbCatalogueElec WHERE CATEGORIE = '" & sCategorie & "' ORDER BY " & sOrderBy, g_connData, adOpenDynamic, adLockOptimistic)

10 iIndex = 1
 
 'Tant que ce n'est pas la fin des enregistrements
Do While Not rstPieces.EOF
 If rstPieces.Fields("PIECE") <> vbNullString And rstPieces.Fields("FABRICANT") <> vbNullString Then
 'Si il y a une recherche à faire
 If sTri <> vbNullString Then
 bDebut = False
 
 'Selon la colonne
 Select Case m_iCol
 'Si c'est la colonne PIECE_GRB
 Case I_COL_PIECES_PIECE_GRB:
 'Si la PIECE_GRB contient la recherche
 If InStr(1, UCase(rstPieces.Fields("PIECE_GRB")), UCase(sTri)) > 0 Then
 'On met la variable à true pour l'ajouter au début
 bDebut = True
 End If
 
 'Si c'est la colonne No. d'item
 Case I_COL_PIECES_NO_ITEM:
 'Si le no. d'item contient la recherche
 If InStr(1, UCase(rstPieces.Fields("PIECE")), UCase(sTri)) > 0 Then
 'On met la variable à true pour l'ajouter au début
 bDebut = True
 End If
 
 'Si c'est la colonne Manufacturier
 Case I_COL_PIECES_MANUFACT:
 'Si le manufacturier contient la recherche
 If InStr(1, UCase(rstPieces.Fields("FABRICANT")), UCase(sTri)) > 0 Then
 'On met la variable à true pour l'ajouter au début
 bDebut = True
 End If
 
 'Si c'est la colonne No. d'item
 Case I_COL_PIECES_DESCR_FR:
 'Si la description française contient la recherche
 If InStr(1, UCase(rstPieces.Fields("DESC_FR")), UCase(sTri)) > 0 Then
 'On met la variable à true pour l'ajouter au début
1  bDebut = True
 End If
 
 'Si c'est la colonne No. d'item
 Case I_COL_PIECES_DESCR_EN:
 'Si la description anglaise contient la recherche
 If InStr(1, UCase(rstPieces.Fields("DESC_EN")), UCase(sTri)) > 0 Then
 'On met la variable à true pour l'ajouter au début
 bDebut = True
 End If
 End Select
 
 If bDebut = True Then
 Set itmPieces = lvwPieces.ListItems.Add(iIndex)
 
 iIndex = iIndex + 1
 Else
 Set itmPieces = lvwPieces.ListItems.Add
 End If
 Else
 Set itmPieces = lvwPieces.ListItems.Add
 End If
 
 'TEMPS
 If Not IsNull(rstPieces.Fields("TEMPS")) Then
 itmPieces.Tag = Trim(rstPieces.Fields("TEMPS"))
 Else
 itmPieces.Tag = vbNullString
 End If
 
 'PIECE_GRB
 If Not IsNull(rstPieces.Fields("PIECE_GRB")) Then
 itmPieces.Text = Trim(rstPieces.Fields("PIECE_GRB"))
 Else
 itmPieces.Text = vbNullString
 End If
 
 'PIECE
 If Not IsNull(rstPieces.Fields("PIECE")) Then
 itmPieces.SubItems(I_COL_PIECES_NO_ITEM) = Trim(rstPieces.Fields("PIECE"))
 Else
 itmPieces.SubItems(I_COL_PIECES_NO_ITEM) = vbNullString
 End If
 
 'FABRICANT
 If Not IsNull(rstPieces.Fields("FABRICANT")) Then
 itmPieces.SubItems(I_COL_PIECES_MANUFACT) = Trim(rstPieces.Fields("FABRICANT"))
 Else
 itmPieces.SubItems(I_COL_PIECES_MANUFACT) = vbNullString
 End If
 
 'DESCR_FR
 If Not IsNull(rstPieces.Fields("DESC_FR")) Then
 itmPieces.SubItems(I_COL_PIECES_DESCR_FR) = Trim(rstPieces.Fields("DESC_FR"))
 Else
 itmPieces.SubItems(I_COL_PIECES_DESCR_FR) = vbNullString
 End If
 
 'DESCR_EN
 If Not IsNull(rstPieces.Fields("DESC_EN")) Then
4 itmPieces.SubItems(I_COL_PIECES_DESCR_EN) = Trim(rstPieces.Fields("DESC_EN"))
4 Else
4 itmPieces.SubItems(I_COL_PIECES_DESCR_EN) = vbNullString
4 End If
4 End If
 
4 Call rstPieces.MoveNext
4 Loop
 
4 Call rstPieces.Close
4 Set rstPieces = Nothing
 
4 Exit Sub

Oups:

4 wOups "frmProjSoumElec", "RemplirListViewPieces", Err, Err.number, Err.Description
End Sub

Private Function TrouverIndexSection(ByVal sSousSection As String) As Integer

 On Error GoTo Oups

 'recherche la section et l'ajouter si elle n'a pas été trouvée
 Dim iCompteur As Integer
 Dim iIndex As Integer
 Dim iTagSection As Integer
 Dim iIDSection As Integer
 Dim iIndexSect As Integer
 Dim bTrouverSect As Boolean
 Dim bTrouverSSect As Boolean
 Dim bTrouverIndexItem As Boolean
 Dim iIndexSSection As Integer
 Dim sTagSousSection As String
  Dim itmSoum As ListItem
 
 'Si la variable sSousSection = PAS DE SOUS-SECTION
  If sSousSection = S_PAS_SOUS_SECTION Then
 'On l'initialise à rien
  sSousSection = vbNullString
 'On met le tag à PAS DE SOUS-SECTION
  sTagSousSection = S_PAS_SOUS_SECTION
  Else
  sTagSousSection = sSousSection
  End If
 
 'Si le listview n'est pas vide
  If lvwSoumission.ListItems.count > 0 Then
 'Pour chaque élément du listview
For iCompteur = 1 To lvwSoumission.ListItems.count
 'Si c'est écrit le nom de la section dans la colonne Piece
 If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) = cmbSections.Text Then
 'La section a été trouvée
 bTrouverSect = True
 
 'On stock l'index de la section
 iIndexSect = iCompteur
 
 'On commence à rechercher la sous-section à l'index suivant
 iCompteur = iCompteur + 1
 
 'Tant que le tag du listItem est égal à l'index de la section
 Do While lvwSoumission.ListItems(iCompteur).Tag = cmbSections.ItemData(cmbSections.ListIndex)
 'Si c'est écrit le nom de la sous-section dans la colonne Description
 If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_DESCR) = sSousSection Then
 'La sous-section a été trouvée
 bTrouverSSect = True
 
 'On stock l'index du premier enregistrement de la section
 iIndex = iCompteur + 1
 
 Exit Do
 End If
 
 iCompteur = iCompteur + 1
 
 'Si le compteur est plus grand que le listItems.Count, il ne faut pas repasser
 'dans la boucle
 If iCompteur > lvwSoumission.ListItems.count Then
 Exit Do
 End If
 Loop
 
 Exit For
 End If
 Next
1  Else
 bTrouverSect = False
 End If
 
If bTrouverSect = False Then
 'Ajoute la section
 
 'Si il y a des enregistrements dans le listview
 If lvwSoumission.ListItems.count > 0 Then
 'Pour chaque élément du listview
 For iCompteur = 1 To lvwSoumission.ListItems.count
 'Si ce n'est pas une section
 If lvwSoumission.ListItems(iCompteur).Tag <> vbNullString Then
 'iTagSection est égal à l'ordre de la section du listitem
 iTagSection = lvwSoumission.ListItems(iCompteur).ListSubItems(I_COL_SOUM_MANUFACT).Tag
 'iIDSection est égal à l'ordre de la section du combo
 iIDSection = cmbSections.ListIndex + 1
 
 'Le premier enregistrement est 2 puisque 1 c'est une section
 If iCompteur = 2 Then
 'Si l'index de la section du combo est plus petit que
 'l'index de la section du ListItem
 If iIDSection < iTagSection Then
 iIndex = 1
 
 Exit For
 End If
 Else
 If iCompteur = lvwSoumission.ListItems.count Then
 'Si l'index de la section du combo est plus grand que l'index
 'de la section du ListItem
 If iIDSection > iTagSection Then
 iIndex = iCompteur + 1
 
 Exit For
 End If
 Else
 If lvwSoumission.ListItems(iCompteur + 1).Tag <> vbNullString Then
 'Si l'index de la section du combo est plus grand que l'index
 'de la section du ListItem et que l'index de la section du combo
 'est plus petit que l'index de la section du ListItem suivant
 If iIDSection > iTagSection And iIDSection < lvwSoumission.ListItems(iCompteur + 1).ListSubItems(I_COL_SOUM_MANUFACT).Tag Then
 iIndex = iCompteur + 1
 
 Exit For
 End If
 Else
 'Si l'index de la section du combo est plus grand que l'index
 'de la section du ListItem et que l'index de la section du combo
 'est plus petit que l'index de 2 ListItem plus loin
 If iIDSection > iTagSection And iIDSection < lvwSoumission.ListItems(iCompteur + 2).ListSubItems(I_COL_SOUM_MANUFACT).Tag Then
 iIndex = iCompteur + 1
 End If
 End If
 End If
 End If
 End If
 Next
Else
 iIndex = 1
End If
 
 'On ajoute la section au bon index
 Set itmSoum = lvwSoumission.ListItems.Add(iIndex)
 
 itmSoum.SubItems(I_COL_SOUM_PIECE) = cmbSections.Text
 itmSoum.ListSubItems(I_COL_SOUM_PIECE).Bold = True
 
itmSoum.SubItems(I_COL_SOUM_MANUFACT) = " "
4 itmSoum.ListSubItems(I_COL_SOUM_MANUFACT).Tag = cmbSections.ListIndex + 1
 
4 Call ValeurParDefaut(itmSoum)
 
4 iIndex = iIndex + 1
 
 'On l'ajoute la sous-section à l'index suivant
4 iIndexSSection = AjouterSousSection(iIndex, sTagSousSection)
 
4 iIndex = iIndexSSection
4 Else
 'Si la sous-section n'a pas été trouvé dans le listview
4 If bTrouverSSect = False Then
 'On l'ajoute à l'index suivant la section
4 iIndexSSection = AjouterSousSection(iIndexSect + 1, sTagSousSection)
 
4 iIndex = iIndexSSection
4 End If
4 End If
 
 'Pour trouver le dernier élément de la sous-section
 
 'Pour chaque élément du listview à partir de l'index du premier élément de la sous-section
4  For iCompteur = iIndex To lvwSoumission.ListItems.count
 'Si on trouve une autre sous-section ou une autre section
4  If lvwSoumission.ListItems(iCompteur).Tag <> cmbSections.ItemData(cmbSections.ListIndex) Or lvwSoumission.ListItems(iCompteur).ListSubItems(I_COL_SOUM_PIECE).Tag <> sTagSousSection Then
4  bTrouverIndexItem = True
 
 'On l'ajoute à cette index
4  iIndex = iCompteur
 
4  Exit For
4  End If
4  Next
 
 'Si la fin de la sous-section n'a pas été, il faut l'ajouter à la fin
4  If bTrouverIndexItem = False Then
50 iIndex = lvwSoumission.ListItems.count + 1
50 End If
 
 TrouverIndexSection = iIndex

 Exit Function

Oups:

 wOups "frmProjSoumElec", "TrouverIndexSection", Err, Err.number, Err.Description
End Function

Private Function AjouterSousSection(ByVal iIndexSection As Integer, ByVal sSousSection As String) As Integer

 On Error GoTo Oups

 'Méthode qui sert à ajouter une sous-section
 Dim itmSoum As ListItem
 Dim iCompteur As Integer
 Dim bTrouverIndexSSection As Boolean
 Dim iIndex As Integer
 Dim sTag As String
 
 If sSousSection = S_PAS_SOUS_SECTION Then
 sSousSection = vbNullString
 sTag = S_PAS_SOUS_SECTION
 Else
 sTag = sSousSection
  End If
 
  If sTag <> S_PAS_SOUS_SECTION Then
 'Pour chaque élément du listview
  For iCompteur = iIndexSection To lvwSoumission.ListItems.count
 'Si le tag du ListItem est différent du IDSection de la section
  If lvwSoumission.ListItems(iCompteur).Tag <> cmbSections.ItemData(cmbSections.ListIndex) Then
  bTrouverIndexSSection = True
 
  iIndex = iCompteur
 
  Exit For
  End If
Next
 
 'Si l'emplacement de la sous-section n'a pas été trouvée
1 If bTrouverIndexSSection = False Then
 'On la place à la fin
 iIndex = lvwSoumission.ListItems.count + 1
 End If
Else
 iIndex = iIndexSection
End If
 
Set itmSoum = lvwSoumission.ListItems.Add(iIndex)
 
Call ValeurParDefaut(itmSoum)
 
 'On met le nom de la sous-section dans la colonne Description
itmSoum.SubItems(I_COL_SOUM_DESCR) = sSousSection
itmSoum.ListSubItems(I_COL_SOUM_DESCR).Bold = True
 
 'On met le nom de la sous-section dans le tag de la colonne Piece
'itmSoum.ListSubItems(I_COL_SOUM_PIECE).Tag = sTag
 
 'On met l'ID de la section dans le tag du listitem
itmSoum.Tag = cmbSections.ItemData(cmbSections.ListIndex)
 
 'On ne peut pas écrire dans le tag si vide
1  itmSoum.SubItems(I_COL_SOUM_MANUFACT) = " "
 
 'Ordre de la section
itmSoum.ListSubItems(I_COL_SOUM_MANUFACT).Tag = cmbSections.ListIndex + 1
 
 AjouterSousSection = iIndex + 1

Exit Function

Oups:

 wOups "frmProjSoumElec", "AjouterSousSection", Err, Err.number, Err.Description
End Function

Private Sub AjouterNegatifDansListView(ByVal dblQuantite As Double, ByVal sSousSection As String)

 On Error GoTo Oups

 Dim iIndex As Integer
 Dim itmSoum As ListItem
 Dim iCompteur As Integer
 Dim iIDSection As Integer
 Dim iTagSection As Integer
 Dim bSelected As Boolean
 Dim iIndexSel As Integer
 Dim dblTempsMec As Double
 Dim lColor As Long
 Dim rstProjet As ADODB.Recordset
  Dim bQteOK As Boolean
  Dim sNoProjet As String
  Dim sPrixList As String
  Dim sEscompte As String
  Dim sPrixNet As String
  Dim sTemps As String
  Dim dblTotalQte As Double

  Set rstProjet = New ADODB.Recordset
 
10 If Right$(txtNoProjSoum.Text, 2) >= 60 And Right$(txtNoProjSoum.Text, 2) <=   Then
1 sNoProjet = Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & m_sLiaison

 If m_bRecherchePiece = True Then
 Call rstProjet.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & sNoProjet & "' AND NumItem = '" & Replace(lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM), "'", "''") & "' AND IDFRS = " & lvwfournisseur.SelectedItem.Tag, g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstProjet.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & sNoProjet & "' AND NumItem = '" & Replace(lvwPieces.SelectedItem.SubItems(I_COL_PIECES_NO_ITEM), "'", "''") & "' AND IDFRS = " & lvwfournisseur.SelectedItem.Tag, g_connData, adOpenDynamic, adLockOptimistic)
 End If
End If

If Not rstProjet.EOF Then
 Do While Not rstProjet.EOF
 dblTotalQte = dblTotalQte + rstProjet.Fields("Qté")

 Call rstProjet.MoveNext
Loop

 If dblTotalQte >= Abs(dblQuantite) Then
 bQteOK = True
 End If
 Else
 Call MsgBox("La pièce n'existe pas dans le projet " & sNoProjet, vbOKOnly, "Erreur")

 Call rstProjet.Close
1  Set rstProjet = Nothing

 Exit Sub
 End If
 
If bQteOK = True Then
 Call rstProjet.MovePrevious

 sPrixList = rstProjet.Fields("Prix_List")
 sEscompte = rstProjet.Fields("Escompte")
 sPrixNet = rstProjet.Fields("Prix_Net")
 sTemps = rstProjet.Fields("Temps")
Else
 If m_bRecherchePiece = True Then
 Call MsgBox("Il n'y a pas assez de " & lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM) & " dans le projet " & sNoProjet & " pour en enlever " & Abs(dblQuantite) & "!", vbOKOnly, "Erreur")
 Else
 Call MsgBox("Il n'y a pas assez de " & lvwPieces.SelectedItem.SubItems(I_COL_PIECES_NO_ITEM) & " dans le projet " & sNoProjet & " pour en enlever " & Abs(dblQuantite) & "!", vbOKOnly, "Erreur")
 End If

Call rstProjet.Close
 Set rstProjet = Nothing

Exit Sub
End If

2  Call rstProjet.Close
Set rstProjet = Nothing
 
30 bSelected = False
 
 'S'il y a des items dans le ListView
If lvwSoumission.ListItems.count > 0 Then
 'Si ce n'est pas le premier qui est sélectionné
 '(le premier est sélectionné par défaut)
 If lvwSoumission.SelectedItem.Index > 1 Then
 bSelected = True

 iIndexSel = lvwSoumission.SelectedItem.Index
 End If
End If

 'si le premier est sélectionné, on le considère comme si aucun n'est sélectionné
If bSelected = False Then
 iIndex = TrouverIndexSection(sSousSection)
Else
 'Sinon, on l'ajoute à l'endroit sélectionné
 iIndex = iIndexSel
End If

3  Set itmSoum = lvwSoumission.ListItems.Add(iIndex)

itmSoum.Checked = True

 'Quantité
3  itmSoum.Text = dblQuantite

If lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_QUOTER) = "Oui" Then
itmSoum.Text = itmSoum.Text & "*"
 itmSoum.ForeColor = COLOR_VERT
 itmSoum.Bold = True
 Else
itmSoum.ForeColor = COLOR_NOIR
4 itmSoum.Bold = False
4 End If

 'On met l'id de la section dans le tag du listItem
4 itmSoum.Tag = cmbSections.ItemData(cmbSections.ListIndex) 'IDSection
    
 'No d'item
4 If m_bRecherchePiece = True Then
4 itmSoum.SubItems(I_COL_SOUM_PIECE) = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM)
4 Else
4 itmSoum.SubItems(I_COL_SOUM_PIECE) = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_NO_ITEM)
4 End If

4 itmSoum.ListSubItems(I_COL_SOUM_PIECE).ForeColor = lColor
 
 'On met le nom de la sous-section dans le tag du no d'item
4 itmSoum.ListSubItems(I_COL_SOUM_PIECE).Tag = sSousSection
 
 'On met la description en francais dans la colonne et la description en anglais
 'dans le tag
4 If m_eLangage = ANGLAIS Then
4  If m_bRecherchePiece = True Then
4  itmSoum.SubItems(I_COL_SOUM_DESCR) = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_DESCR_EN)
4  itmSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_DESCR_FR)
4  Else
4  itmSoum.SubItems(I_COL_SOUM_DESCR) = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_DESCR_EN)
4  itmSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_DESCR_FR)
4  End If
4  Else
50 If m_bRecherchePiece = True Then
itmSoum.SubItems(I_COL_SOUM_DESCR) = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_DESCR_FR)
 itmSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_DESCR_EN)
 Else
 itmSoum.SubItems(I_COL_SOUM_DESCR) = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_DESCR_FR)
 itmSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_DESCR_EN)
 End If
 End If

 itmSoum.ListSubItems(I_COL_SOUM_DESCR).ForeColor = lColor
 
 'On met le nom du fabricant dans la colonne et l'ordre de la section dans le tag
 If m_bRecherchePiece = True Then
 itmSoum.SubItems(I_COL_SOUM_MANUFACT) = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_MANUFACT)
 Else
5  itmSoum.SubItems(I_COL_SOUM_MANUFACT) = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_MANUFACT)
5  End If

5  itmSoum.ListSubItems(I_COL_SOUM_MANUFACT).Tag = cmbSections.ListIndex + 1 'Ordre de la section
 
5  itmSoum.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = lColor

 'Prix listé
5  If Trim$(sPrixList) = vbNullString Then
5  itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion("0", MODE_ARGENT, 4)
5  Else
5  itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(sPrixList, MODE_ARGENT, 4)
60 itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = sPrixList
60 End If
 
  itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = lColor
 
 'Si il y a un prix net, on le met l'escompte et le prix net sinon, on prend le prix
 'spécial pour mettre dans le prix net
  If Trim$(sEscompte) <> vbNullString Then
  itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = sEscompte
  Else
  itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion("0", MODE_POURCENT)
  End If
 
  itmSoum.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = lColor

  If Trim$(sPrixNet) <> vbNullString Then
  itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(sPrixNet, MODE_ARGENT, 4)
  Else
6  itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion("0", MODE_ARGENT, 4)
6  End If

6  itmSoum.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = lColor

6  itmSoum.SubItems(I_COL_SOUM_DISTRIB) = lvwfournisseur.SelectedItem.Text
6  itmSoum.ListSubItems(I_COL_SOUM_DISTRIB).Tag = lvwfournisseur.SelectedItem.Tag
 
6  itmSoum.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = lColor
 
 'Temps
6  itmSoum.SubItems(I_COL_SOUM_TEMPS) = sTemps

6  itmSoum.ListSubItems(I_COL_SOUM_TEMPS).ForeColor = lColor
 
 'Si le temps n'est pas vide
70 If Trim$(itmSoum.SubItems(I_COL_SOUM_TEMPS)) <> vbNullString Then
 'On calcul le temps * quantité pour la colonne montage
  itmSoum.SubItems(I_COL_SOUM_MONTAGE) = CDbl(Replace(itmSoum.SubItems(I_COL_SOUM_TEMPS), ".", ",")) * CDbl(Replace(itmSoum.Text, "*", vbNullString))
  Else
  itmSoum.SubItems(I_COL_SOUM_MONTAGE) = vbNullString
  End If
 
  itmSoum.ListSubItems(I_COL_SOUM_MONTAGE).ForeColor = lColor
 
 'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
  itmSoum.SubItems(I_COL_SOUM_TOTAL) = Conversion(Round(Replace(itmSoum.Text, "*", vbNullString) * itmSoum.SubItems(I_COL_SOUM_PRIX_NET) * CSng(m_sProfit), 2), MODE_ARGENT, 2)
 
  itmSoum.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = lColor
 
 'Pour le profit, c'est le prix total - (prix net * quantité)
  itmSoum.SubItems(I_COL_SOUM_PROFIT) = Conversion(CStr(Round(itmSoum.SubItems(I_COL_SOUM_TOTAL) - (itmSoum.SubItems(I_COL_SOUM_PRIX_NET) * Replace(itmSoum.Text, "*", vbNullString)), 2)), MODE_ARGENT)

  itmSoum.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = lColor

  If itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = vbNullString Then
  itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = " "
   End If

   Call CalculerTempsFabrication

7  Exit Sub

Oups:

7  wOups "frmProjSoumElec", "AjouterNegatifDansListView", Err, Err.number, Err.Description
End Sub

Private Sub AjouterDansListViewSoumission(ByVal dblQuantite As Double, ByVal sSousSection As String)

 On Error GoTo Oups

 Dim rstConfig As ADODB.Recordset
 Dim iIndex As Integer
 Dim iCompteur As Integer
 Dim iIDSection As Integer
 Dim iTagSection As Integer
 Dim iIndexSel As Integer
 Dim itmSoum As ListItem
 Dim bSelected As Boolean
 Dim dblTempsMec As Double
 Dim sDistrib As String
  Dim sTauxUSA As String
  Dim sTauxSPA As String
  Dim lColor As Long
 
  bSelected = False
 
 'S'il y a des items dans le ListView
  If lvwSoumission.ListItems.count > 0 Then
 'Si ce n'est pas le premier qui est sélectionné
 '(le premier est sélectionné par défaut)
  If lvwSoumission.SelectedItem.Index > 1 Then
  bSelected = True
 
  iIndexSel = lvwSoumission.SelectedItem.Index
End If
End If
 
 'Si le premier est sélectionné, on le considère comme si aucun n'est sélectionné
If bSelected = False Then
 iIndex = TrouverIndexSection(sSousSection)
Else
 'Sinon, on l'ajoute à l'endroit sélectionné
 iIndex = iIndexSel
End If
 
Set rstConfig = New ADODB.Recordset

Call rstConfig.Open("SELECT TauxAmericain, TauxEspagnol FROM GrbConfig", g_connData, adOpenForwardOnly, adLockReadOnly)
 
sTauxUSA = rstConfig.Fields("TauxAmericain")
sTauxSPA = rstConfig.Fields("TauxEspagnol")

Call rstConfig.Close
1  Set rstConfig = Nothing
 
Set itmSoum = lvwSoumission.ListItems.Add(iIndex)
 
 itmSoum.Checked = True
 
 'Quantité
itmSoum.Text = dblQuantite
 
 If lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_QUOTER) = "Oui" Then
 itmSoum.Text = itmSoum.Text & "*"
 itmSoum.ForeColor = COLOR_VERT
1  itmSoum.Bold = True
 Else
 itmSoum.ForeColor = COLOR_NOIR
 itmSoum.Bold = False
End If
 
If lvwfournisseur.SelectedItem.Text = "CHOISIR ULTÉRIEUREMENT" Then
 lColor = COLOR_MAGENTA
Else
 lColor = COLOR_NOIR
End If

 'On met l'id de la section dans le tag du listItem
itmSoum.Tag = cmbSections.ItemData(cmbSections.ListIndex)
    
 'No d'item
If m_bRecherchePiece = True Then
 itmSoum.SubItems(I_COL_SOUM_PIECE) = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM)
2  Else
 itmSoum.SubItems(I_COL_SOUM_PIECE) = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_NO_ITEM)
2  End If

itmSoum.ListSubItems(I_COL_SOUM_PIECE).ForeColor = lColor
 
 'On met le nom de la sous-section dans le tag du no d'item
2  itmSoum.ListSubItems(I_COL_SOUM_PIECE).Tag = sSousSection
 
 'On met la description en francais dans la colonne et la description en anglais
 'dans le tag
If m_eLangage = ANGLAIS Then
If m_bRecherchePiece = True Then
 itmSoum.SubItems(I_COL_SOUM_DESCR) = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_DESCR_EN)
 itmSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_DESCR_FR)
3 Else
 itmSoum.SubItems(I_COL_SOUM_DESCR) = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_DESCR_EN)
 itmSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_DESCR_FR)
 End If
Else
 If m_bRecherchePiece = True Then
 itmSoum.SubItems(I_COL_SOUM_DESCR) = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_DESCR_FR)
 itmSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_DESCR_EN)
 Else
 itmSoum.SubItems(I_COL_SOUM_DESCR) = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_DESCR_FR)
 itmSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_DESCR_EN)
End If
End If

3  itmSoum.ListSubItems(I_COL_SOUM_DESCR).ForeColor = lColor
 
 'On met le nom du fabricant dans la colonne et l'ordre de la section dans le tag
If m_bRecherchePiece = True Then
itmSoum.SubItems(I_COL_SOUM_MANUFACT) = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_MANUFACT)
Else
 itmSoum.SubItems(I_COL_SOUM_MANUFACT) = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_MANUFACT)
 End If

40 itmSoum.ListSubItems(I_COL_SOUM_MANUFACT).Tag = cmbSections.ListIndex + 1 'Ordre de la section
 
itmSoum.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = lColor

 'Prix listé
4 If Trim$(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_LIST)) = vbNullString Then
4 itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion("0", MODE_ARGENT, 4)
4 Else
4 If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "USA" Then
4 itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_LIST)) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
4 Else
4 If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "SPA" Then
4 itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_LIST)) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
4 Else
4 itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_LIST), MODE_ARGENT, 4)
4  End If
4  End If
4  End If

4  itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PRIX_LIST).Tag
 
4  itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = lColor
 
 'Si il y a un prix net, on le met l'escompte et le prix net sinon, on prend le prix
 'spécial pour mettre dans le prix net
4  If Trim$(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_NET)) <> vbNullString Then
4  If Trim$(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_ESCOMPTE)) <> vbNullString Then
4  itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_ESCOMPTE), MODE_POURCENT)
50 Else
itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion(0, MODE_POURCENT)
 End If

 If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "USA" Then
 itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_NET)) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
 Else
 If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "SPA" Then
 itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_NET)) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
 Else
 itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_NET), MODE_ARGENT, 4)
 End If
 End If
5  Else
5  If Trim$(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_SP)) <> vbNullString Then
5  If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "USA" Then
5  itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_SP)) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
5  Else
5  If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "SPA" Then
5  itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_SP)) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
5  Else
60 itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_SP), MODE_ARGENT, 4)
  End If
  End If
  Else
  itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion("0", MODE_POURCENT)
  itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion("0", MODE_ARGENT, 4)
  End If
  End If
 
  itmSoum.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = lColor
  itmSoum.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = lColor
 
 'On met le fournisseur dans la colonne et l'id dans le tag
  If lvwfournisseur.SelectedItem.Text = "CHOISIR ULTÉRIEUREMENT" Then
  sDistrib = vbNullString
6  Else
6  sDistrib = lvwfournisseur.SelectedItem.Text
6  End If

6  itmSoum.SubItems(I_COL_SOUM_DISTRIB) = sDistrib
6  itmSoum.ListSubItems(I_COL_SOUM_DISTRIB).Tag = lvwfournisseur.SelectedItem.Tag
 
6  itmSoum.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = lColor
 
 'Temps
6  If m_bRecherchePiece = True Then
6  itmSoum.SubItems(I_COL_SOUM_TEMPS) = Replace(lvwPieceTrouve.SelectedItem.Tag, ".", ",")
70 Else
  itmSoum.SubItems(I_COL_SOUM_TEMPS) = Replace(lvwPieces.SelectedItem.Tag, ".", ",")
  End If

  itmSoum.ListSubItems(I_COL_SOUM_TEMPS).ForeColor = lColor
 
 'Si le temps n'est pas vide
  If Trim$(itmSoum.SubItems(I_COL_SOUM_TEMPS)) <> vbNullString Then
 'On calcul le temps * quantité pour la colonne montage
  itmSoum.SubItems(I_COL_SOUM_MONTAGE) = CDbl(itmSoum.SubItems(I_COL_SOUM_TEMPS)) * CDbl(Replace(itmSoum.Text, "*", vbNullString))
  Else
  itmSoum.SubItems(I_COL_SOUM_MONTAGE) = vbNullString
  End If
 
  itmSoum.ListSubItems(I_COL_SOUM_MONTAGE).ForeColor = lColor
 
 'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
  itmSoum.SubItems(I_COL_SOUM_TOTAL) = Conversion(Round(Replace(itmSoum.Text, "*", vbNullString) * itmSoum.SubItems(I_COL_SOUM_PRIX_NET) * CSng(m_sProfit), 2), MODE_ARGENT, 2)
 
  itmSoum.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = lColor

   itmSoum.ListSubItems(I_COL_SOUM_TOTAL).Tag = lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag
 
 'Pour le profit, c'est le prix total - (prix net * quantité)
   itmSoum.SubItems(I_COL_SOUM_PROFIT) = Conversion(CStr(Round(itmSoum.SubItems(I_COL_SOUM_TOTAL) - (itmSoum.SubItems(I_COL_SOUM_PRIX_NET) * Replace(itmSoum.Text, "*", vbNullString)), 2)), MODE_ARGENT)

7  itmSoum.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = lColor

7  If itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = vbNullString Then
7  itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = " "
7  End If

7  Call CalculerTempsFabrication

7  Call itmSoum.EnsureVisible

80 Exit Sub

Oups:

80 wOups "frmProjSoumElec", "AjouterDansListViewSoumission", Err, Err.number, Err.Description
End Sub

Private Sub CalculerTempsFabrication()

 On Error GoTo Oups

 Dim dblTempsFab As Double
 Dim iCompteur As Integer

 'Pour chaque élément du listView
 For iCompteur = 1 To lvwSoumission.ListItems.count
 If Trim$(lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_MONTAGE)) <> vbNullString Then
 'On additionne le temps
 dblTempsFab = dblTempsFab + CDbl(Replace(Trim$(lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_MONTAGE)), ".", ","))
 End If
 Next
 
 m_sTempsFabrication = Replace(dblTempsFab / 10, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "CalculerTempsFabrication", Err, Err.number, Err.Description
End Sub

Private Function VerifierEmplacement(ByVal iIndexSelection As Integer) As Boolean

 On Error GoTo Oups

 'Vérifie si l'emplacement pour ajouter une pièce est valide
 Dim itmSoum As ListItem
 
 Set itmSoum = lvwSoumission.ListItems(iIndexSelection)
 
 If itmSoum.Tag = vbNullString Then
 Set itmSoum = lvwSoumission.ListItems(iIndexSelection - 1)
 End If
 
 'Si la section est correcte
 If itmSoum.Tag = cmbSections.ItemData(cmbSections.ListIndex) Then
 VerifierEmplacement = True
 Else
 VerifierEmplacement = False
 End If

  Exit Function

Oups:

  wOups "frmProjSoumElec", "VerifierEmplacement", Err, Err.number, Err.Description
End Function

Private Sub ValeurParDefaut(ByVal itmSoumission As ListItem)

 On Error GoTo Oups

 'Méthode pour mettre une valeur par défaut dans quelques colonnes de lvwSoumission.
 'Si ces colonnes sont vides, elles restent blanches lors de la sélection
 If m_bDroitPrix = True Then
 itmSoumission.SubItems(I_COL_SOUM_PRIX_LIST) = " "
 itmSoumission.SubItems(I_COL_SOUM_ESCOMPTE) = " "
 itmSoumission.SubItems(I_COL_SOUM_PRIX_NET) = " "
 itmSoumission.SubItems(I_COL_SOUM_TOTAL) = " "
 itmSoumission.SubItems(I_COL_SOUM_PROFIT) = " "
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "ValeurParDefaut", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListViewProjSoum(ByVal sNoProjSoum As String)

 On Error GoTo Oups

 'Remplis les pièces de la soumission avec la BD
 Dim rstProjSoum As ADODB.Recordset
 Dim rstSection As ADODB.Recordset
 Dim rstFRS As ADODB.Recordset
 Dim itmProjSoum As ListItem
 Dim bPremierEnr As Boolean
 Dim iOrdreSection As Integer
 Dim sSousSection As String
 Dim sSection As String
 Dim lColor As Long
 Dim bBold As Boolean

  Set rstProjSoum = New ADODB.Recordset
  Set rstSection = New ADODB.Recordset
  Set rstFRS = New ADODB.Recordset
 
  Call lvwSoumission.ListItems.Clear
 
  bPremierEnr = True
 
  If m_eType = TYPE_PROJET Then
  Call rstProjSoum.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & sNoProjSoum & "' ORDER BY CInt(OrdreSection), NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
  Else
Call rstProjSoum.Open("SELECT * FROM GrbSoumission_Pieces WHERE IDSoumission = '" & sNoProjSoum & "' ORDER BY CInt(OrdreSection), NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
End If
 
If m_eLangage = ANGLAIS Then
 sSection = "NomSectionEN"
Else
 sSection = "NomSectionFR"
End If
 
Do While Not rstProjSoum.EOF
 Set itmProjSoum = lvwSoumission.ListItems.Add
 
 'Si c'est le premier enregistrement, il faut ajouter la section et la sous-section
 If bPremierEnr = True Then
 iOrdreSection = rstProjSoum.Fields("OrdreSection")
 sSousSection = rstProjSoum.Fields("SousSection")
 
 'Pour avoir le nom de la section
 Call rstSection.Open("SELECT " & sSection & " FROM GrbSoumProjSectionElec WHERE IDSection = " & rstProjSoum.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
 
 'Ajout du nom de la section
 If Not IsNull(rstSection.Fields(sSection)) Then
 itmProjSoum.SubItems(I_COL_SOUM_PIECE) = rstSection.Fields(sSection)
 Else
 itmProjSoum.SubItems(I_COL_SOUM_PIECE) = vbNullString
 End If
 
 itmProjSoum.ListSubItems(I_COL_SOUM_PIECE).Bold = True
 
1  Call ValeurParDefaut(itmProjSoum)
 
 Call rstSection.Close
 
 Set itmProjSoum = lvwSoumission.ListItems.Add
 
 'Ajout du nom de la sous-section
 If sSousSection = S_PAS_SOUS_SECTION Then
 itmProjSoum.SubItems(I_COL_SOUM_DESCR) = vbNullString
 Else
 itmProjSoum.SubItems(I_COL_SOUM_DESCR) = sSousSection
 End If
 
 'Le tag ne peut pas être remplis si la colonne est vide
 itmProjSoum.SubItems(I_COL_SOUM_MANUFACT) = " "
 
 itmProjSoum.ListSubItems(I_COL_SOUM_MANUFACT).Tag = iOrdreSection
 
 itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Bold = True
 
 itmProjSoum.Tag = rstProjSoum.Fields("IDSection")
 
 Call ValeurParDefaut(itmProjSoum)
 
 Set itmProjSoum = lvwSoumission.ListItems.Add
 
 bPremierEnr = False
Else
 'Si c'est pas le premier enregistrement, il faut vérifier avec l'ancienne section
 If iOrdreSection <> rstProjSoum.Fields("OrdreSection") Then
 iOrdreSection = rstProjSoum.Fields("OrdreSection")
 
 Call rstSection.Open("SELECT " & sSection & " FROM GrbSoumProjSectionElec WHERE IDSection = " & rstProjSoum.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
 
 If Not IsNull(rstSection.Fields(sSection)) Then
 itmProjSoum.SubItems(I_COL_SOUM_PIECE) = rstSection.Fields(sSection)
 Else
 itmProjSoum.SubItems(I_COL_SOUM_PIECE) = vbNullString
 End If
 
 itmProjSoum.ListSubItems(I_COL_SOUM_PIECE).Bold = True
 
 Call ValeurParDefaut(itmProjSoum)
 
 Call rstSection.Close
 
 Set itmProjSoum = lvwSoumission.ListItems.Add
 
 sSousSection = rstProjSoum.Fields("SousSection")
 
 If sSousSection = S_PAS_SOUS_SECTION Then
 itmProjSoum.SubItems(I_COL_SOUM_DESCR) = vbNullString
 Else
 itmProjSoum.SubItems(I_COL_SOUM_DESCR) = rstProjSoum.Fields("SousSection")
 End If
 
 'Le tag ne peut pas être remplis si la colonne est vide
 itmProjSoum.SubItems(I_COL_SOUM_MANUFACT) = " "
 
 itmProjSoum.ListSubItems(I_COL_SOUM_MANUFACT).Tag = iOrdreSection
 
 itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Bold = True
 
 Call ValeurParDefaut(itmProjSoum)
 
 itmProjSoum.Tag = rstProjSoum.Fields("IDSection")
 
 Set itmProjSoum = lvwSoumission.ListItems.Add
 Else
 'il faut vérifier avec l'ancienne sous-section
 If sSousSection <> rstProjSoum.Fields("SousSection") Then
4 sSousSection = rstProjSoum.Fields("SousSection")
 
4 If sSousSection = S_PAS_SOUS_SECTION Then
4 itmProjSoum.SubItems(I_COL_SOUM_DESCR) = vbNullString
4 Else
4 itmProjSoum.SubItems(I_COL_SOUM_DESCR) = sSousSection
4 End If
 
4 itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Bold = True
 
4 Call ValeurParDefaut(itmProjSoum)
 
4 itmProjSoum.Tag = rstProjSoum.Fields("IDSection")
 
 'Le tag ne peut pas être remplis si la colonne est vide
4 itmProjSoum.SubItems(I_COL_SOUM_MANUFACT) = " "
 
4 itmProjSoum.ListSubItems(I_COL_SOUM_MANUFACT).Tag = iOrdreSection
 
4  Set itmProjSoum = lvwSoumission.ListItems.Add
4  End If
4  End If
4  End If

4  If rstProjSoum.Fields("PieceExtraChargeable") = True Then
4  lColor = COLOR_BLEU
4  bBold = True
4  Else
50 If rstProjSoum.Fields("PieceExtraNonChargeable") = True Then
 lColor = COLOR_ROSE
 bBold = True
 Else
 If rstProjSoum.Fields("CommandeAnnulée") = True Then
 lColor = COLOR_VERT_FORET
 bBold = True
 Else
 If rstProjSoum.Fields("Retour") = True Then
 lColor = COLOR_ROUGE
 bBold = False
 Else
5  If rstProjSoum.Fields("Commandé") = True Then
5  lColor = COLOR_ORANGE 'COLOR_ORANGE
5  bBold = False
5  Else
5  If rstProjSoum.Fields("Recu") = True Then
5  lColor = COLOR_GRIS 'Gris
5  bBold = False
5  Else
60 If rstProjSoum.Fields("IDFRS") = 0 And rstProjSoum.Fields("NumItem") <> "Texte" And rstProjSoum.Fields("NumItem") <> "Text" Then
  lColor = COLOR_MAGENTA
  bBold = False
  Else
  If rstProjSoum.Fields("MatérielInutile") = True Then
  lColor = COLOR_BRUN
  bBold = False
  Else
  lColor = COLOR_NOIR
  bBold = False
  End If
  End If
6  End If
6  End If
6  End If
6  End If
6  End If
6  End If

 'On met l'ID de la section dans le tag
6  itmProjSoum.Tag = rstProjSoum.Fields("IDSection")
 
6  If rstProjSoum.Fields("Visible") = True Then
70 itmProjSoum.Checked = True
  Else
  itmProjSoum.Checked = False
  End If
 
 'Quantité
  If Not IsNull(rstProjSoum.Fields("Qté")) Then
  itmProjSoum.Text = rstProjSoum.Fields("Qté")
  Else
  itmProjSoum.Text = vbNullString
  End If
 
 'On met la quantité en vert avec un astérix si il est quoté
  If rstProjSoum.Fields("Quoté") = True Then
  itmProjSoum.Text = itmProjSoum.Text & "*"
  itmProjSoum.ForeColor = COLOR_VERT
   itmProjSoum.Bold = True
   Else
7  itmProjSoum.ForeColor = COLOR_NOIR
7  itmProjSoum.Bold = False
7  End If

 'Facturation
7  If m_eType = TYPE_PROJET Then
7  If g_bModificationProjetsElec = True Then
7  If Not IsNull(rstProjSoum.Fields("Facturation")) Then
80 If Trim(rstProjSoum.Fields("Facturation")) <> "" Then
  itmProjSoum.SubItems(I_COL_SOUM_FACTURATION) = rstProjSoum.Fields("Facturation")
  Else
  itmProjSoum.SubItems(I_COL_SOUM_FACTURATION) = " "
  End If
  Else
  itmProjSoum.SubItems(I_COL_SOUM_FACTURATION) = " "
  End If
  End If
  End If

 'Numéro d'item
  If Not IsNull(rstProjSoum.Fields("NumItem")) Then
  If rstProjSoum.Fields("NumItem") = "Texte" Or rstProjSoum.Fields("NumItem") = "Text" Then
   If m_eLangage = ANGLAIS Then
   itmProjSoum.SubItems(I_COL_SOUM_PIECE) = "Text"
   Else
   itmProjSoum.SubItems(I_COL_SOUM_PIECE) = "Texte"
8  End If
8  Else
8  itmProjSoum.SubItems(I_COL_SOUM_PIECE) = rstProjSoum.Fields("NumItem")
8  End If
90 Else
  itmProjSoum.SubItems(I_COL_SOUM_PIECE) = vbNullString
  End If
 
  itmProjSoum.ListSubItems(I_COL_SOUM_PIECE).ForeColor = lColor
  itmProjSoum.ListSubItems(I_COL_SOUM_PIECE).Bold = bBold
 
 'On met le nom de la sous-section dans le tag du numéro d'item
  itmProjSoum.ListSubItems(I_COL_SOUM_PIECE).Tag = rstProjSoum.Fields("SousSection")

  If itmProjSoum.SubItems(I_COL_SOUM_PIECE) = "Texte" Or itmProjSoum.SubItems(I_COL_SOUM_PIECE) = "Text" Then
  itmProjSoum.SubItems(I_COL_SOUM_DESCR) = rstProjSoum.Fields("DESC_FR")
  Else
  If m_eLangage = ANGLAIS Then
 'Description en anglais
  If Not IsNull(rstProjSoum.Fields("DESC_EN")) Then
  itmProjSoum.SubItems(I_COL_SOUM_DESCR) = rstProjSoum.Fields("DESC_EN")
 Else
   itmProjSoum.SubItems(I_COL_SOUM_DESCR) = vbNullString
 End If
 
 'On met la description en francais dans le tag de la description en anglais
   If Not IsNull(rstProjSoum.Fields("DESC_FR")) Then
 itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = rstProjSoum.Fields("DESC_FR")
   Else
 itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = vbNullString
9  End If
 Else
 'Description en francais
 If Not IsNull(rstProjSoum.Fields("DESC_FR")) Then
 itmProjSoum.SubItems(I_COL_SOUM_DESCR) = rstProjSoum.Fields("DESC_FR")
 Else
 itmProjSoum.SubItems(I_COL_SOUM_DESCR) = vbNullString
 End If
 
 'On met la description en anglais dans le tag de la description en francais
 If Not IsNull(rstProjSoum.Fields("DESC_EN")) Then
 itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = rstProjSoum.Fields("DESC_EN")
 Else
 itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = vbNullString
 End If
1 End If
10  End If
 
10  itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).ForeColor = lColor
10  itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Bold = bBold

 'Fabricant
10  If Not IsNull(rstProjSoum.Fields("Manufact")) Then
10  itmProjSoum.SubItems(I_COL_SOUM_MANUFACT) = rstProjSoum.Fields("Manufact")
10  Else
10  itmProjSoum.SubItems(I_COL_SOUM_MANUFACT) = vbNullString
10  End If
 
110 itmProjSoum.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = lColor

11 itmProjSoum.ListSubItems(I_COL_SOUM_MANUFACT).Bold = bBold
 
 'On met l'ordre de la section dans le tag du fabricant
1 itmProjSoum.ListSubItems(I_COL_SOUM_MANUFACT).Tag = iOrdreSection
 
 'Prix listé
1 If m_bDroitPrix = True Then
1 If Not IsNull(rstProjSoum.Fields("PRIX_LIST")) Then
1 If rstProjSoum.Fields("PRIX_LIST") <> "" Then
1 itmProjSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(rstProjSoum.Fields("PRIX_LIST"), MODE_ARGENT, 4)
1 Else
1 itmProjSoum.SubItems(I_COL_SOUM_PRIX_LIST) = " "
1 End If
1 Else
1 itmProjSoum.SubItems(I_COL_SOUM_PRIX_LIST) = " "
1 End If

1 itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = lColor
 itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).Bold = bBold
 
1 itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = rstProjSoum.Fields("PrixOrigine")

 'Escompte
 If Trim(rstProjSoum.Fields("Escompte")) <> vbNullString Then
1 itmProjSoum.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion(rstProjSoum.Fields("Escompte"), MODE_POURCENT)
 Else
11  itmProjSoum.SubItems(I_COL_SOUM_ESCOMPTE) = " "
 End If
 
 itmProjSoum.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = lColor
 
1 itmProjSoum.ListSubItems(I_COL_SOUM_ESCOMPTE).Bold = bBold
 
1 If Not IsNull(rstProjSoum.Fields("PRIX_NET")) Then
1 If rstProjSoum.Fields("PRIX_NET") <> "" Then
1 itmProjSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(rstProjSoum.Fields("PRIX_NET"), MODE_ARGENT, 4)
1 Else
1 itmProjSoum.SubItems(I_COL_SOUM_PRIX_NET) = " "
1 End If
1 Else
1 itmProjSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion("", MODE_ARGENT, 4)
1 End If
 
1 itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = lColor
 
1 itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_NET).Bold = bBold
 
1 If m_eType = TYPE_PROJET Then
1 itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_NET).Tag = rstProjSoum.Fields("DateRéception")
1 End If
 
 'Fournisseur
1 If Not IsNull(rstProjSoum.Fields("IDFRS")) And rstProjSoum.Fields("IDFRS") > 0 Then
1 If itmProjSoum.SubItems(I_COL_SOUM_PIECE) <> "Texte" And itmProjSoum.SubItems(I_COL_SOUM_PIECE) <> "Text" Then
1 Call rstFRS.Open("SELECT NomFournisseur FROM GrbFournisseur WHERE IDFRS = " & rstProjSoum.Fields("IDFRS"), g_connData, adOpenDynamic, adLockOptimistic)
 
 'On affiche le nom dans la colonne
1 itmProjSoum.SubItems(I_COL_SOUM_DISTRIB) = rstFRS.Fields("NomFournisseur")
 
1 itmProjSoum.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = lColor
1 itmProjSoum.ListSubItems(I_COL_SOUM_DISTRIB).Bold = bBold
 
 'On affiche l'Id dans le tag
1 itmProjSoum.ListSubItems(I_COL_SOUM_DISTRIB).Tag = rstProjSoum.Fields("IDFRS")
 
1 Call rstFRS.Close
1 End If
1 Else
1 itmProjSoum.SubItems(I_COL_SOUM_DISTRIB) = vbNullString
1 itmProjSoum.ListSubItems(I_COL_SOUM_DISTRIB).Tag = 0
1 End If
 
 'Temps
1 If Not IsNull(rstProjSoum.Fields("Temps")) Then
1 itmProjSoum.SubItems(I_COL_SOUM_TEMPS) = rstProjSoum.Fields("Temps")
1 Else
1 itmProjSoum.SubItems(I_COL_SOUM_TEMPS) = vbNullString
1 End If

1 itmProjSoum.ListSubItems(I_COL_SOUM_TEMPS).ForeColor = lColor

1 itmProjSoum.ListSubItems(I_COL_SOUM_TEMPS).Bold = bBold
 
 'Montage
1 If Not IsNull(rstProjSoum.Fields("Temps_total")) Then
1 itmProjSoum.SubItems(I_COL_SOUM_MONTAGE) = rstProjSoum.Fields("Temps_total")
1 Else
1 itmProjSoum.SubItems(I_COL_SOUM_MONTAGE) = vbNullString
14 End If
 
14 itmProjSoum.ListSubItems(I_COL_SOUM_MONTAGE).ForeColor = lColor
 
14 itmProjSoum.ListSubItems(I_COL_SOUM_MONTAGE).Bold = bBold
 
 'Prix total
14 If Trim(rstProjSoum.Fields("Prix_total")) <> vbNullString Then
14 If IsNumeric(rstProjSoum.Fields("Prix_Total")) Then
14 itmProjSoum.SubItems(I_COL_SOUM_TOTAL) = Conversion(Round(rstProjSoum.Fields("Prix_total"), 2), MODE_ARGENT)
14 Else
14 itmProjSoum.SubItems(I_COL_SOUM_TOTAL) = " "
14 End If
14 Else
14 itmProjSoum.SubItems(I_COL_SOUM_TOTAL) = " "
14  End If
 
14  itmProjSoum.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = lColor

14  itmProjSoum.ListSubItems(I_COL_SOUM_TOTAL).Bold = bBold
 
14  itmProjSoum.ListSubItems(I_COL_SOUM_TOTAL).Tag = rstProjSoum.Fields("Devise")
 
 'Profit
14  If Trim(rstProjSoum.Fields("Profit_argent")) <> vbNullString Then
14  itmProjSoum.SubItems(I_COL_SOUM_PROFIT) = Conversion(Round(rstProjSoum.Fields("Profit_Argent"), 2), MODE_ARGENT)
14  Else
14  itmProjSoum.SubItems(I_COL_SOUM_PROFIT) = " "
150 End If

15 If m_eType = TYPE_PROJET Then
 If rstProjSoum.Fields("PieceExtraChargeable") = True Or rstProjSoum.Fields("PieceExtraNonChargeable") = True Then
 itmProjSoum.ListSubItems(I_COL_SOUM_PROFIT).Tag = "EXTRA"
 End If
 End If
 
 itmProjSoum.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = lColor

 itmProjSoum.ListSubItems(I_COL_SOUM_PROFIT).Bold = bBold

 If Not IsNull(rstProjSoum.Fields("Commentaire")) Then
 itmProjSoum.SubItems(I_COL_SOUM_COMMENTAIRE) = rstProjSoum.Fields("Commentaire")
 Else
 itmProjSoum.SubItems(I_COL_SOUM_COMMENTAIRE) = vbNullString
15  End If

15  itmProjSoum.ListSubItems(I_COL_SOUM_COMMENTAIRE).ForeColor = lColor

15  itmProjSoum.ListSubItems(I_COL_SOUM_COMMENTAIRE).Bold = bBold

15  If m_eType = TYPE_PROJET Then
15  If Not IsNull(rstProjSoum.Fields("ID")) Then
15  itmProjSoum.SubItems(I_COL_SOUM_ID) = rstProjSoum.Fields("ID")
15  Else
15  itmProjSoum.SubItems(I_COL_SOUM_ID) = vbNullString
160 End If

 itmProjSoum.ListSubItems(I_COL_SOUM_ID).ForeColor = lColor

 itmProjSoum.ListSubItems(I_COL_SOUM_ID).Bold = bBold

 If Not IsNull(rstProjSoum.Fields("DateCommande")) Then
 If Trim(rstProjSoum.Fields("DateCommande")) <> "" Then
 itmProjSoum.SubItems(I_COL_SOUM_DATE_COMMANDE) = rstProjSoum.Fields("DateCommande")
 Else
 itmProjSoum.SubItems(I_COL_SOUM_DATE_COMMANDE) = " "
 End If
 Else
 itmProjSoum.SubItems(I_COL_SOUM_DATE_COMMANDE) = " "
 End If

16  itmProjSoum.ListSubItems(I_COL_SOUM_DATE_COMMANDE).ForeColor = lColor

16  itmProjSoum.ListSubItems(I_COL_SOUM_DATE_COMMANDE).Bold = bBold

16  itmProjSoum.ListSubItems(I_COL_SOUM_DATE_COMMANDE).Tag = rstProjSoum.Fields("NoRetour")

16  If Not IsNull(rstProjSoum.Fields("DateRequise")) Then
16  If Trim(rstProjSoum.Fields("DateRequise")) <> "" Then
16  itmProjSoum.SubItems(I_COL_SOUM_DATE_REQUISE) = rstProjSoum.Fields("DateRequise")
16  Else
16  itmProjSoum.SubItems(I_COL_SOUM_DATE_REQUISE) = " "
170 End If
 Else
 itmProjSoum.SubItems(I_COL_SOUM_DATE_REQUISE) = " "
 End If

 itmProjSoum.ListSubItems(I_COL_SOUM_DATE_REQUISE).ForeColor = lColor

 itmProjSoum.ListSubItems(I_COL_SOUM_DATE_REQUISE).Bold = bBold

 itmProjSoum.ListSubItems(I_COL_SOUM_DATE_REQUISE).Tag = rstProjSoum.Fields("DateRetour")

 If Not IsNull(rstProjSoum.Fields("NomCommande")) Then
 itmProjSoum.SubItems(I_COL_SOUM_NOM_COMMANDE) = rstProjSoum.Fields("NomCommande")
 Else
 itmProjSoum.SubItems(I_COL_SOUM_NOM_COMMANDE) = vbNullString
 End If

1   itmProjSoum.ListSubItems(I_COL_SOUM_NOM_COMMANDE).ForeColor = lColor

1   itmProjSoum.ListSubItems(I_COL_SOUM_NOM_COMMANDE).Bold = bBold

17  If Not IsNull(rstProjSoum.Fields("NoSéquentiel")) Then
17  itmProjSoum.SubItems(I_COL_SOUM_NO_SEQUENTIEL) = rstProjSoum.Fields("NoSéquentiel")
17  Else
17  itmProjSoum.SubItems(I_COL_SOUM_NO_SEQUENTIEL) = vbNullString
17  End If

17  itmProjSoum.ListSubItems(I_COL_SOUM_NO_SEQUENTIEL).ForeColor = lColor

180 itmProjSoum.ListSubItems(I_COL_SOUM_NO_SEQUENTIEL).Bold = bBold

 If Not IsNull(rstProjSoum.Fields("Provenance")) Then
 If Trim(rstProjSoum.Fields("Provenance")) <> vbNullString Then
 itmProjSoum.SubItems(I_COL_SOUM_PROVENANCE) = Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 3) & "-" & rstProjSoum.Fields("Provenance")
 Else
 itmProjSoum.SubItems(I_COL_SOUM_PROVENANCE) = vbNullString
 End If
 Else
 itmProjSoum.SubItems(I_COL_SOUM_PROVENANCE) = vbNullString
 End If

 itmProjSoum.ListSubItems(I_COL_SOUM_PROVENANCE).ForeColor = lColor
 itmProjSoum.ListSubItems(I_COL_SOUM_PROVENANCE).Bold = bBold
1   Else
1   If Not IsNull(rstProjSoum.Fields("Provenance")) Then
1   If Trim(rstProjSoum.Fields("Provenance")) <> vbNullString Then
1   itmProjSoum.SubItems(I_COL_SOUMISSION_PROV) = Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 3) & "-" & rstProjSoum.Fields("Provenance")
18  Else
18  itmProjSoum.SubItems(I_COL_SOUMISSION_PROV) = vbNullString
18  End If
18  Else
190 itmProjSoum.SubItems(I_COL_SOUMISSION_PROV) = vbNullString
 End If

1  itmProjSoum.ListSubItems(I_COL_SOUMISSION_PROV).ForeColor = lColor
1  itmProjSoum.ListSubItems(I_COL_SOUMISSION_PROV).Bold = bBold
1  End If
1  Else
 'Fournisseur
1  If Not IsNull(rstProjSoum.Fields("IDFRS")) And rstProjSoum.Fields("IDFRS") > 0 Then
1  If itmProjSoum.SubItems(I_COL_SOUM_SP_PIECE) <> "Texte" And itmProjSoum.SubItems(I_COL_SOUM_SP_PIECE) <> "Text" Then
1  Call rstFRS.Open("SELECT NomFournisseur FROM GrbFournisseur WHERE IDFRS = " & rstProjSoum.Fields("IDFRS"), g_connData, adOpenDynamic, adLockOptimistic)
 
 'On affiche le nom dans la colonne
1  itmProjSoum.SubItems(I_COL_SOUM_SP_DISTRIB) = rstFRS.Fields("NomFournisseur")
 
1  itmProjSoum.ListSubItems(I_COL_SOUM_SP_DISTRIB).ForeColor = lColor

1  itmProjSoum.ListSubItems(I_COL_SOUM_SP_DISTRIB).Bold = bBold
 
 'On affiche l'Id dans le tag
 itmProjSoum.ListSubItems(I_COL_SOUM_SP_DISTRIB).Tag = rstProjSoum.Fields("IDFRS")
 
1   Call rstFRS.Close
 End If
1   Else
 itmProjSoum.SubItems(I_COL_SOUM_SP_DISTRIB) = vbNullString
1   itmProjSoum.ListSubItems(I_COL_SOUM_SP_DISTRIB).Tag = 0
 End If
 
 'Temps
19  If Not IsNull(rstProjSoum.Fields("Temps")) Then
200 itmProjSoum.SubItems(I_COL_SOUM_SP_TEMPS) = rstProjSoum.Fields("Temps")
 Else
 itmProjSoum.SubItems(I_COL_SOUM_SP_TEMPS) = vbNullString
 End If
 
 itmProjSoum.ListSubItems(I_COL_SOUM_SP_TEMPS).ForeColor = lColor

 itmProjSoum.ListSubItems(I_COL_SOUM_SP_TEMPS).Bold = bBold
 
 'Montage
 If Not IsNull(rstProjSoum.Fields("Temps_total")) Then
 itmProjSoum.SubItems(I_COL_SOUM_SP_MONTAGE) = rstProjSoum.Fields("Temps_total")
 Else
 itmProjSoum.SubItems(I_COL_SOUM_SP_MONTAGE) = vbNullString
 End If

 itmProjSoum.ListSubItems(I_COL_SOUM_SP_MONTAGE).ForeColor = lColor

20  itmProjSoum.ListSubItems(I_COL_SOUM_SP_MONTAGE).Bold = bBold

20  If Not IsNull(rstProjSoum.Fields("Commentaire")) Then
20  itmProjSoum.SubItems(I_COL_SOUM_SP_COMMENTAIRE) = rstProjSoum.Fields("Commentaire")
20  Else
20  itmProjSoum.SubItems(I_COL_SOUM_SP_COMMENTAIRE) = vbNullString
20  End If

20  itmProjSoum.ListSubItems(I_COL_SOUM_SP_COMMENTAIRE).ForeColor = lColor

20  itmProjSoum.ListSubItems(I_COL_SOUM_SP_COMMENTAIRE).Bold = bBold

2 If m_eType = TYPE_PROJET Then
2 If Not IsNull(rstProjSoum.Fields("ID")) Then
 itmProjSoum.SubItems(I_COL_SOUM_SP_ID) = rstProjSoum.Fields("ID")
 Else
 itmProjSoum.SubItems(I_COL_SOUM_SP_ID) = vbNullString
 End If

 itmProjSoum.ListSubItems(I_COL_SOUM_SP_ID).ForeColor = lColor

 itmProjSoum.ListSubItems(I_COL_SOUM_SP_ID).Bold = bBold

 If Not IsNull(rstProjSoum.Fields("DateCommande")) Then
 If Trim(rstProjSoum.Fields("DateCommande")) <> "" Then
 itmProjSoum.SubItems(I_COL_SOUM_SP_DATE_COMMANDE) = rstProjSoum.Fields("DateCommande")
 Else
 itmProjSoum.SubItems(I_COL_SOUM_SP_DATE_COMMANDE) = " "
 End If
 Else
 itmProjSoum.SubItems(I_COL_SOUM_SP_DATE_COMMANDE) = " "
 End If

 itmProjSoum.ListSubItems(I_COL_SOUM_SP_DATE_COMMANDE).ForeColor = lColor

 itmProjSoum.ListSubItems(I_COL_SOUM_SP_DATE_COMMANDE).Bold = bBold

21  itmProjSoum.ListSubItems(I_COL_SOUM_SP_DATE_COMMANDE).Tag = rstProjSoum.Fields("DateCommande")

 If Not IsNull(rstProjSoum.Fields("DateRequise")) Then
 If Trim(rstProjSoum.Fields("DateRequise")) <> "" Then
2 itmProjSoum.SubItems(I_COL_SOUM_SP_DATE_REQUISE) = rstProjSoum.Fields("DateRequise")
2 Else
2 itmProjSoum.SubItems(I_COL_SOUM_SP_DATE_REQUISE) = ""
2 End If
2 Else
2 itmProjSoum.SubItems(I_COL_SOUM_SP_DATE_REQUISE) = " "
2 End If

2 itmProjSoum.ListSubItems(I_COL_SOUM_SP_DATE_REQUISE).ForeColor = lColor

2 itmProjSoum.ListSubItems(I_COL_SOUM_SP_DATE_REQUISE).Bold = bBold

2 itmProjSoum.ListSubItems(I_COL_SOUM_SP_DATE_REQUISE).Tag = rstProjSoum.Fields("DateRetour")

2 itmProjSoum.SubItems(I_COL_SOUM_SP_NOM_COMMANDE) = rstProjSoum.Fields("NomCommande")

2 itmProjSoum.ListSubItems(I_COL_SOUM_SP_NOM_COMMANDE).ForeColor = lColor

2 itmProjSoum.ListSubItems(I_COL_SOUM_SP_NOM_COMMANDE).Bold = bBold

2 itmProjSoum.SubItems(I_COL_SOUM_SP_NO_SEQUENTIEL) = rstProjSoum.Fields("NoSéquentiel")

2 itmProjSoum.ListSubItems(I_COL_SOUM_SP_NO_SEQUENTIEL).ForeColor = lColor

2 itmProjSoum.ListSubItems(I_COL_SOUM_SP_NO_SEQUENTIEL).Bold = bBold

2 If Not IsNull(rstProjSoum.Fields("Provenance")) Then
2 If Trim(rstProjSoum.Fields("Provenance")) <> "" Then
2 itmProjSoum.SubItems(I_COL_SOUM_SP_PROVENANCE) = Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 3) & "-" & rstProjSoum.Fields("Provenance")
2 Else
2 itmProjSoum.SubItems(I_COL_SOUM_SP_PROVENANCE) = ""
2 End If
2 Else
2 itmProjSoum.SubItems(I_COL_SOUM_SP_PROVENANCE) = ""
2 End If

2 itmProjSoum.ListSubItems(I_COL_SOUM_SP_PROVENANCE).ForeColor = lColor
2 itmProjSoum.ListSubItems(I_COL_SOUM_SP_PROVENANCE).Bold = bBold
2 Else
2 If Not IsNull(rstProjSoum.Fields("Provenance")) Then
2 If Trim(rstProjSoum.Fields("Provenance")) <> "" Then
2 itmProjSoum.SubItems(I_COL_SOUMISSION_SP_PROV) = Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 3) & "-" & rstProjSoum.Fields("Provenance")
2 Else
2 itmProjSoum.SubItems(I_COL_SOUMISSION_SP_PROV) = ""
2 End If
2 Else
2 itmProjSoum.SubItems(I_COL_SOUMISSION_SP_PROV) = ""
2 End If

2 itmProjSoum.ListSubItems(I_COL_SOUMISSION_SP_PROV).ForeColor = lColor
2 itmProjSoum.ListSubItems(I_COL_SOUMISSION_SP_PROV).Bold = bBold
24 End If
24 End If
 
24 Call rstProjSoum.MoveNext
 
24 Call lvwSoumission.Refresh
24 Loop

24 If lvwSoumission.ListItems.count > 0 Then
24 Call Deselect

24 lvwSoumission.ListItems(1).Selected = True
24 End If

24 Call CalculerPrix
 
24 Call rstProjSoum.Close
24  Set rstProjSoum = Nothing

24  Set rstFRS = Nothing
24  Set rstSection = Nothing

24  Exit Sub

Oups:

248woups"frmProjSoumElec", "RemplirListViewProjSoum", Err, Erl, sNoProjSoum)
End Sub

Private Sub RemplirListViewSoumissionProjet(ByVal sNoProjet As String)

 On Error GoTo Oups

 'Remplis les pièces de la soumission avec la BD
 Dim rstProjSoum As ADODB.Recordset
 Dim rstSection As ADODB.Recordset
 Dim rstFRS As ADODB.Recordset
 Dim itmProjSoum As ListItem
 Dim bPremierEnr As Boolean
 Dim iOrdreSection As Integer
 Dim sSousSection As String
 Dim sSection As String
 Dim lColor As Long

 Set rstProjSoum = New ADODB.Recordset
  Set rstSection = New ADODB.Recordset
  Set rstFRS = New ADODB.Recordset
 
  Call lvwSoumission.ListItems.Clear
 
  bPremierEnr = True
 
  Call rstProjSoum.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & sNoProjet & "' ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
 
  If m_eLangage = ANGLAIS Then
  sSection = "NomSectionEN"
  Else
sSection = "NomSectionFR"
End If
 
Do While Not rstProjSoum.EOF
 Set itmProjSoum = lvwSoumission.ListItems.Add
 
 'Si c'est le premier enregistrement, il faut ajouter la section et la sous-section
 If bPremierEnr = True Then
 iOrdreSection = rstProjSoum.Fields("OrdreSection")
 sSousSection = rstProjSoum.Fields("SousSection")
 
 'Pour avoir le nom de la section
 Call rstSection.Open("SELECT " & sSection & " FROM GrbSoumProjSectionElec WHERE IDSection = " & rstProjSoum.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
 
 'Ajout du nom de la section
 If Not IsNull(rstSection.Fields(sSection)) Then
 itmProjSoum.SubItems(I_COL_SOUM_PIECE) = rstSection.Fields(sSection)
 Else
 itmProjSoum.SubItems(I_COL_SOUM_PIECE) = vbNullString
 End If
 
 itmProjSoum.ListSubItems(I_COL_SOUM_PIECE).Bold = True
 
 Call ValeurParDefaut(itmProjSoum)
 
 Call rstSection.Close
 
 Set itmProjSoum = lvwSoumission.ListItems.Add
 
 'Ajout du nom de la sous-section
 If sSousSection = S_PAS_SOUS_SECTION Then
 itmProjSoum.SubItems(I_COL_SOUM_DESCR) = vbNullString
1  Else
 itmProjSoum.SubItems(I_COL_SOUM_DESCR) = sSousSection
 End If
 
 'Le tag ne peut pas être remplis si la colonne est vide
 itmProjSoum.SubItems(I_COL_SOUM_MANUFACT) = " "
 
 itmProjSoum.ListSubItems(I_COL_SOUM_MANUFACT).Tag = iOrdreSection
 
 itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Bold = True
 
 itmProjSoum.Tag = rstProjSoum.Fields("IDSection")
 
 Call ValeurParDefaut(itmProjSoum)
 
 Set itmProjSoum = lvwSoumission.ListItems.Add
 
 bPremierEnr = False
 Else
 'Si c'est pas le premier enregistrement, il faut vérifier avec l'ancienne section
 If iOrdreSection <> rstProjSoum.Fields("OrdreSection") Then
 iOrdreSection = rstProjSoum.Fields("OrdreSection")
 
 Call rstSection.Open("SELECT " & sSection & " FROM GrbSoumProjSectionElec WHERE IDSection = " & rstProjSoum.Fields("IDSection"), g_connData, adOpenDynamic, adLockOptimistic)
 
 If Not IsNull(rstSection.Fields(sSection)) Then
 itmProjSoum.SubItems(I_COL_SOUM_PIECE) = rstSection.Fields(sSection)
 Else
 itmProjSoum.SubItems(I_COL_SOUM_PIECE) = vbNullString
 End If
 
 itmProjSoum.ListSubItems(I_COL_SOUM_PIECE).Bold = True
 
 Call ValeurParDefaut(itmProjSoum)
 
 Call rstSection.Close
 
 Set itmProjSoum = lvwSoumission.ListItems.Add
 
 sSousSection = rstProjSoum.Fields("SousSection")
 
 If sSousSection = S_PAS_SOUS_SECTION Then
 itmProjSoum.SubItems(I_COL_SOUM_DESCR) = vbNullString
 Else
 itmProjSoum.SubItems(I_COL_SOUM_DESCR) = rstProjSoum.Fields("SousSection")
 End If
 
 'Le tag ne peut pas être remplis si la colonne est vide
 itmProjSoum.SubItems(I_COL_SOUM_MANUFACT) = " "
 
 itmProjSoum.ListSubItems(I_COL_SOUM_MANUFACT).Tag = iOrdreSection
 
 itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Bold = True
 
 Call ValeurParDefaut(itmProjSoum)
 
 itmProjSoum.Tag = rstProjSoum.Fields("IDSection")
 
 Set itmProjSoum = lvwSoumission.ListItems.Add
 Else
 'il faut vérifier avec l'ancienne sous-section
 If sSousSection <> rstProjSoum.Fields("SousSection") Then
 sSousSection = rstProjSoum.Fields("SousSection")
 
 If sSousSection = S_PAS_SOUS_SECTION Then
 itmProjSoum.SubItems(I_COL_SOUM_DESCR) = vbNullString
 Else
 itmProjSoum.SubItems(I_COL_SOUM_DESCR) = sSousSection
4 End If
 
4 itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Bold = True
 
4 Call ValeurParDefaut(itmProjSoum)
 
4 itmProjSoum.Tag = rstProjSoum.Fields("IDSection")
 
 'Le tag ne peut pas être remplis si la colonne est vide
4 itmProjSoum.SubItems(I_COL_SOUM_MANUFACT) = " "
 
4 itmProjSoum.ListSubItems(I_COL_SOUM_MANUFACT).Tag = iOrdreSection
 
4 Set itmProjSoum = lvwSoumission.ListItems.Add
4 End If
4 End If
4 End If

4 If rstProjSoum.Fields("IDFRS") = 0 And rstProjSoum.Fields("NumItem") <> "Texte" And rstProjSoum.Fields("NumItem") <> "Text" Then
4  lColor = COLOR_MAGENTA
4  Else
4  lColor = COLOR_NOIR
4  End If

 'On met l'ID de la section dans le tag
4  itmProjSoum.Tag = rstProjSoum.Fields("IDSection")
 
4  If rstProjSoum.Fields("Visible") = True Then
4  itmProjSoum.Checked = True
4  Else
50 itmProjSoum.Checked = False
5 End If
 
 'Quantité
 If Not IsNull(rstProjSoum.Fields("Qté")) Then
 itmProjSoum.Text = rstProjSoum.Fields("Qté")
 Else
 itmProjSoum.Text = vbNullString
 End If
 
 'On met la quantité en vert avec un astérix si il est quoté
 If rstProjSoum.Fields("Quoté") = True Then
 itmProjSoum.Text = itmProjSoum.Text & "*"
 itmProjSoum.ForeColor = COLOR_VERT
 itmProjSoum.Bold = True
 Else
5  itmProjSoum.ForeColor = COLOR_NOIR
5  itmProjSoum.Bold = False
5  End If

 'Numéro d'item
5  If Not IsNull(rstProjSoum.Fields("NumItem")) Then
5  If rstProjSoum.Fields("NumItem") = "Texte" Or rstProjSoum.Fields("NumItem") = "Text" Then
5  If m_eLangage = ANGLAIS Then
5  itmProjSoum.SubItems(I_COL_SOUM_PIECE) = "Text"
5  Else
60 itmProjSoum.SubItems(I_COL_SOUM_PIECE) = "Texte"
  End If
  Else
  itmProjSoum.SubItems(I_COL_SOUM_PIECE) = rstProjSoum.Fields("NumItem")
  End If
  Else
  itmProjSoum.SubItems(I_COL_SOUM_PIECE) = vbNullString
  End If
 
  itmProjSoum.ListSubItems(I_COL_SOUM_PIECE).ForeColor = lColor
 
 'On met le nom de la sous-section dans le tag du numéro d'item
  itmProjSoum.ListSubItems(I_COL_SOUM_PIECE).Tag = rstProjSoum.Fields("SousSection")

  If itmProjSoum.SubItems(I_COL_SOUM_PIECE) = "Texte" Or itmProjSoum.SubItems(I_COL_SOUM_PIECE) = "Text" Then
  itmProjSoum.SubItems(I_COL_SOUM_DESCR) = rstProjSoum.Fields("DESC_FR")
6  Else
6  If m_eLangage = ANGLAIS Then
 'Description en anglais
6  If Not IsNull(rstProjSoum.Fields("DESC_EN")) Then
6  itmProjSoum.SubItems(I_COL_SOUM_DESCR) = rstProjSoum.Fields("DESC_EN")
6  Else
6  itmProjSoum.SubItems(I_COL_SOUM_DESCR) = vbNullString
6  End If
 
 'On met la description en francais dans le tag de la description en anglais
6  If Not IsNull(rstProjSoum.Fields("DESC_FR")) Then
70 itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = rstProjSoum.Fields("DESC_FR")
  Else
  itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = vbNullString
  End If
  Else
 'Description en francais
  If Not IsNull(rstProjSoum.Fields("DESC_FR")) Then
  itmProjSoum.SubItems(I_COL_SOUM_DESCR) = rstProjSoum.Fields("DESC_FR")
  Else
  itmProjSoum.SubItems(I_COL_SOUM_DESCR) = vbNullString
  End If
 
 'On met la description en anglais dans le tag de la description en francais
  If Not IsNull(rstProjSoum.Fields("DESC_EN")) Then
  itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = rstProjSoum.Fields("DESC_EN")
   Else
   itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = vbNullString
7  End If
7  End If
7  End If
 
7  itmProjSoum.ListSubItems(I_COL_SOUM_DESCR).ForeColor = lColor

 'Fabricant
7  If Not IsNull(rstProjSoum.Fields("Manufact")) Then
7  itmProjSoum.SubItems(I_COL_SOUM_MANUFACT) = rstProjSoum.Fields("Manufact")
80 Else
  itmProjSoum.SubItems(I_COL_SOUM_MANUFACT) = vbNullString
  End If
 
  itmProjSoum.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = lColor
 
 'On met l'ordre de la section dans le tag du fabricant
  itmProjSoum.ListSubItems(I_COL_SOUM_MANUFACT).Tag = iOrdreSection
 
 'Prix listé
  If Trim(rstProjSoum.Fields("Prix_List")) <> vbNullString Then
  itmProjSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(rstProjSoum.Fields("Prix_list"), MODE_ARGENT, 4)
  Else
  itmProjSoum.SubItems(I_COL_SOUM_PRIX_LIST) = " "
  End If
 
  itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = lColor
 
  itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = rstProjSoum.Fields("PrixOrigine")
 
 'Escompte
   If Trim(rstProjSoum.Fields("Escompte")) <> vbNullString Then
   itmProjSoum.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion(rstProjSoum.Fields("Escompte"), MODE_POURCENT)
   Else
   itmProjSoum.SubItems(I_COL_SOUM_ESCOMPTE) = " "
8  End If
 
8  itmProjSoum.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = lColor
 
 'Prix net
8  If Trim(rstProjSoum.Fields("Prix_net")) <> vbNullString Then
8  itmProjSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(rstProjSoum.Fields("Prix_net"), MODE_ARGENT, 4)
90 Else
  itmProjSoum.SubItems(I_COL_SOUM_PRIX_NET) = " "
  End If
 
  itmProjSoum.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = lColor
 
 'Fournisseur
  If Not IsNull(rstProjSoum.Fields("IDFRS")) And rstProjSoum.Fields("IDFRS") > 0 Then
  If itmProjSoum.SubItems(I_COL_SOUM_PIECE) <> "Texte" And itmProjSoum.SubItems(I_COL_SOUM_PIECE) <> "Text" Then
  Call rstFRS.Open("SELECT NomFournisseur FROM GrbFournisseur WHERE IDFRS = " & rstProjSoum.Fields("IDFRS"), g_connData, adOpenDynamic, adLockOptimistic)
 
 'On affiche le nom dans la colonne
  itmProjSoum.SubItems(I_COL_SOUM_DISTRIB) = rstFRS.Fields("NomFournisseur")
 
  itmProjSoum.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = lColor
 
 'On affiche l'Id dans le tag
  itmProjSoum.ListSubItems(I_COL_SOUM_DISTRIB).Tag = rstProjSoum.Fields("IDFRS")
 
  Call rstFRS.Close
  End If
 Else
   itmProjSoum.SubItems(I_COL_SOUM_DISTRIB) = vbNullString
 itmProjSoum.ListSubItems(I_COL_SOUM_DISTRIB).Tag = 0
   End If
 
 'Temps
 If Not IsNull(rstProjSoum.Fields("Temps")) Then
   itmProjSoum.SubItems(I_COL_SOUM_TEMPS) = rstProjSoum.Fields("Temps")
 Else
9  itmProjSoum.SubItems(I_COL_SOUM_TEMPS) = vbNullString
 End If
 
10 itmProjSoum.ListSubItems(I_COL_SOUM_TEMPS).ForeColor = lColor
 
 'Montage
1 If Not IsNull(rstProjSoum.Fields("Temps_total")) Then
1 itmProjSoum.SubItems(I_COL_SOUM_MONTAGE) = rstProjSoum.Fields("Temps_total")
 Else
1 itmProjSoum.SubItems(I_COL_SOUM_MONTAGE) = vbNullString
 End If
 
1 itmProjSoum.ListSubItems(I_COL_SOUM_MONTAGE).ForeColor = lColor
 
 'Prix total
 If Trim(rstProjSoum.Fields("Prix_total")) <> vbNullString Then
1 itmProjSoum.SubItems(I_COL_SOUM_TOTAL) = Conversion(Round(rstProjSoum.Fields("Prix_total"), 2), MODE_ARGENT)
 Else
1 itmProjSoum.SubItems(I_COL_SOUM_TOTAL) = " "
10  End If
 
10  itmProjSoum.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = lColor
 
 'Profit
10  If Trim(rstProjSoum.Fields("Profit_argent")) <> vbNullString Then
10  itmProjSoum.SubItems(I_COL_SOUM_PROFIT) = Conversion(Round(rstProjSoum.Fields("Profit_Argent"), 2), MODE_ARGENT)
10  Else
10  itmProjSoum.SubItems(I_COL_SOUM_PROFIT) = " "
10  End If
 
10  itmProjSoum.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = lColor

110 If Not IsNull(rstProjSoum.Fields("Commentaire")) Then
11 itmProjSoum.SubItems(I_COL_SOUM_COMMENTAIRE) = rstProjSoum.Fields("Commentaire")
1 Else
1 itmProjSoum.SubItems(I_COL_SOUM_COMMENTAIRE) = vbNullString
1 End If

1 itmProjSoum.ListSubItems(I_COL_SOUM_COMMENTAIRE).ForeColor = lColor
 
1 Call rstProjSoum.MoveNext
 
1 Call lvwSoumission.Refresh
11 Loop
 
11 Call rstProjSoum.Close
11 Set rstProjSoum = Nothing

11 Set rstFRS = Nothing
11  Set rstSection = Nothing

11  Exit Sub

Oups:

1 wOups "frmProjSoumElec", "RemplirListSoumissionProjet", Err, Err.number, Err.Description
End Sub

Private Sub CalculerPrix()

 On Error GoTo Oups

 'Méthode pour calculer le prix
 Dim dblPrixPieces As Double
 Dim dblPrixTotal As Double
 Dim dblCommission As Double
 Dim dblTotalTemps As Double
 Dim dblProfit As Double
 Dim dblTotalManuel As Double
 Dim dblTotalImprevue As Double
 Dim dblGrandTotal As Double
 Dim dblTotalDessin As Double
 Dim dblTotalFabrication As Double
  Dim dblTotalAssemblage As Double
  Dim dblTotalProgInterface As Double
  Dim dblTotalProgAutomate As Double
  Dim dblTotalProgRobot As Double
  Dim dblTotalVision As Double
  Dim dblTotalTest As Double
  Dim dblTotalInstallation As Double
  Dim dblTotalMiseService As Double
10 Dim dblTotalFormation As Double
Dim dblTotalGestion As Double
Dim dblTotalShipping As Double
Dim dblHebergement As Double
Dim dblRepas As Double
Dim dblTransport As Double
Dim dblUniteMobile As Double
Dim dblPrixEmballage As Double
Dim dblTotalResteTemps As Double
Dim bDemande As Boolean
Dim iNbrePersonne As Integer
Dim iCompteur As Integer
 
 'Si ce n'est pas en mode affichage
1  If m_bModeAffichage = False Then
 'Pour chaque élément du listview
 For iCompteur = 1 To lvwSoumission.ListItems.count
 'Si ce n'est pas une section
 If lvwSoumission.ListItems(iCompteur).Tag <> vbNullString Then
 'Si ce n'est pas une sous-section
 If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> vbNullString And lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> "Texte" And lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> "Text" Then
 If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_DISTRIB) <> vbNullString Then
 'On additionne le prix total
 
 If IsNumeric(lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_TOTAL)) And IsNumeric(lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PROFIT)) Then
 dblPrixPieces = dblPrixPieces + lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_TOTAL) - lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PROFIT)
1  Else
 Call MsgBox("La pièce " & lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) & " a un prix non numérique!", vbOKOnly, "Erreur")
 End If
 
 'On additionne le profit
 If IsNumeric(Trim$(lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PROFIT))) = True Then
 dblProfit = dblProfit + lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PROFIT)
 End If
 Else
 bDemande = True
 End If
 End If
 End If
 Next
 
 'Total des temps
 dblTotalDessin = CDbl(m_sTempsDessin) * CDbl(m_sTauxDessin)

If m_bSansTemps = False Then
 dblTotalFabrication = CDbl(m_sTempsFabrication) * CDbl(m_sTauxFabrication)
Else
 dblTotalFabrication = 0
End If
 
 dblTotalAssemblage = CDbl(m_sTempsAssemblage) * CDbl(m_sTauxAssemblage)
dblTotalProgInterface = CDbl(m_sTempsProgInterface) * CDbl(m_sTauxProgInterface)
 dblTotalProgAutomate = CDbl(m_sTempsProgAutomate) * CDbl(m_sTauxProgAutomate)
dblTotalProgRobot = CDbl(m_sTempsProgRobot) * CDbl(m_sTauxProgRobot)
3 dblTotalVision = CDbl(m_sTempsVision) * CDbl(m_sTauxVision)
 dblTotalTest = CDbl(m_sTempsTest) * CDbl(m_sTauxTest)
 dblTotalInstallation = CDbl(m_sTempsInstallation) * CDbl(m_sTauxInstallation)
 dblTotalMiseService = CDbl(m_sTempsMiseService) * CDbl(m_sTauxMiseService)
 dblTotalFormation = CDbl(m_sTempsFormation) * CDbl(m_sTauxFormation)
 dblTotalGestion = CDbl(m_sTempsGestion) * CDbl(m_sTauxGestion)
 dblTotalShipping = CDbl(m_sTempsShipping) * CDbl(m_sTauxShipping)

 dblTotalTemps = dblTotalDessin + _
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
 
 If m_eType = TYPE_PROJET Then
 dblHebergement = 0
 dblRepas = 0
 dblTransport = 0
 dblUniteMobile = 0
Else
 iNbrePersonne = Int(m_sNbrePersonne)
 
 Do While iNbrePersonne > 0
 If iNbrePersonne >= 2 Then
 dblHebergement = dblHebergement + CDbl(m_sTempsHebergement) * CDbl(m_sTauxHebergement2)
 
 iNbrePersonne = iNbrePersonne - 2
 Else
4 dblHebergement = dblHebergement + CDbl(m_sTempsHebergement) * CDbl(m_sTauxHebergement1)
 
4 iNbrePersonne = iNbrePersonne - 1
4 End If
4 Loop
 
4 dblRepas = CDbl(m_sTempsRepas) * CDbl(m_sTauxRepas) * CDbl(m_sNbrePersonne)
4 dblTransport = CDbl(m_sTempsTransport) * CDbl(m_sTauxTransport)
4 dblUniteMobile = CDbl(m_sTempsUniteMobile) * CDbl(m_sTauxUniteMobile)
4 End If

4 If IsNumeric(m_sPrixEmballage) Then
4 dblPrixEmballage = CDbl(m_sPrixEmballage)
4 Else
4  dblPrixEmballage = 0
4  End If
 
4  dblTotalResteTemps = dblHebergement + dblRepas + dblTransport + dblUniteMobile + dblPrixEmballage
  
4  If IsNumeric(txtPrixManuel.Text) Then
4  dblTotalManuel = CDbl(txtPrixManuel.Text)
4  Else
4  dblTotalManuel = 0
4  End If
 
50 dblTotalImprevue = (dblPrixPieces + dblProfit) * CDbl(m_sImprevue)
 
5 dblPrixTotal = dblPrixPieces + dblProfit + dblTotalTemps + dblTotalImprevue + dblTotalManuel + dblTotalResteTemps
 
 'Le calcul de la commission est sur les manuels (Nbre de page * prix par pages * nbre de copies) + (prix des pièces * pourcentage d'imprévus) + total des temps
 dblCommission = dblPrixTotal * CDbl(m_sCommission)
 
 'Le prix total est le calcul des manuels (Nbre de page * prix par pages * nbre de copies) + (prix des pièces * pourcentage d'imprévus) + total des temps + total de la commission
 dblGrandTotal = dblPrixTotal + dblCommission
 
 'Format monétaires avec 2 chiffres après la virgule
 txtTotalPieces.Text = Conversion(CStr(Round(dblPrixPieces, 2)), MODE_ARGENT)
 txtTotalTemps.Text = Conversion(CStr(Round(dblTotalTemps, 2)), MODE_ARGENT)
 txtPrixTotal.Text = Conversion(CStr(Round(dblGrandTotal, 2)), MODE_ARGENT)
 
 If bDemande = True Then
 txtPrixTotal.ForeColor = COLOR_JAUNE
 Else
 txtPrixTotal.ForeColor = COLOR_ROUGE
 End If

5  txtImprevus.Text = Conversion(CStr(Round(dblTotalImprevue, 2)), MODE_ARGENT)
5  txtCommission.Text = Conversion(CStr(Round(dblCommission, 2)), MODE_ARGENT)
5  txtProfit.Text = Conversion(CStr(Round(dblProfit, 2)), MODE_ARGENT)
5  Else
5  For iCompteur = 1 To lvwSoumission.ListItems.count
 'Si ce n'est pas une section
5  If lvwSoumission.ListItems(iCompteur).Tag <> vbNullString Then
 'Si ce n'est pas une sous-section
5  If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> vbNullString And lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> "Texte" And lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> "Text" Then
5  If m_bDroitPrix = True Then
60 If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_DISTRIB) = vbNullString Then
  bDemande = True
 
  Exit For
  End If
  End If
  End If
  End If
  Next

  If bDemande = True Then
  txtPrixTotal.ForeColor = COLOR_JAUNE
  Else
  txtPrixTotal.ForeColor = COLOR_ROUGE
6  End If
6  End If

6  Exit Sub

Oups:

6  wOups "frmProjSoumElec", "CalculerPrix", Err, Err.number, Err.Description
End Sub

Private Sub CalculerTempsFabricationRecordset(ByVal sNoProjSoum As String)

 On Error GoTo Oups

 Dim rstProjet As ADODB.Recordset
 Dim rstPiece As ADODB.Recordset
 Dim dblTempsFab As Double

 'Ouverture des tables
 Set rstProjet = New ADODB.Recordset
 Set rstPiece = New ADODB.Recordset

 Call rstProjet.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)

 Call rstPiece.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet ='" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)

 'Pour chaque enregistrement du recordset
 Do While Not rstPiece.EOF
 'Si le temps total n'est pas vide
 If Trim(rstPiece.Fields("Temps_total")) <> vbNullString Then
 'On additionne le temps
 dblTempsFab = dblTempsFab + CDbl(Replace(Trim(rstPiece.Fields("Temps_total")), ".", ","))
  End If

  Call rstPiece.MoveNext
  Loop
 
  rstProjet.Fields("TempsFabrication") = Replace(dblTempsFab / 10, ".", ",")

  Call rstProjet.Update

  Call rstPiece.Close
  Set rstPiece = Nothing

  Call rstProjet.Close
10 Set rstProjet = Nothing

Exit Sub

Oups:

wOups "frmProjSoumElec", "CalculerTempsFabricationRecordset", Err, Err.number, Err.Description
End Sub

Private Sub CalculerTotalRecordset(ByVal sNoProjSoum As String)

 On Error GoTo Oups

 'Méthode pour calculer le prix
 Dim rstProjSoum As ADODB.Recordset
 Dim rstPiece As ADODB.Recordset
 Dim rstPunch As ADODB.Recordset
 Dim dblTotalDessin As Double
 Dim dblTotalFabrication As Double
 Dim dblTotalAssemblage As Double
 Dim dblTotalProgInterface As Double
 Dim dblTotalProgAutomate As Double
 Dim dblTotalProgRobot As Double
 Dim dblTotalVision As Double
  Dim dblTotalTest As Double
  Dim dblTotalInstallation As Double
  Dim dblTotalMiseService As Double
  Dim dblTotalFormation As Double
  Dim dblTotalGestion As Double
  Dim dblTotalShipping As Double
  Dim dblHebergement As Double
  Dim dblRepas As Double
10 Dim dblTransport As Double
Dim dblUniteMobile As Double
Dim dblPrixEmballage As Double
Dim dblTotalResteTemps As Double
Dim dblPrixPieces As Double
Dim dblPrixTotal As Double
Dim dblCommission As Double
Dim dblTotalTemps As Double
Dim dblProfit As Double
Dim dblTotalManuel As Double
Dim dblTotalPieceImprevue As Double
Dim dblGrandTotal As Double
1  Dim sDateDebut As String
Dim sDateFin As String
 Dim sTotal As String
Dim sFilterNoProjet As String

 Set rstProjSoum = New ADODB.Recordset

If m_eType = TYPE_PROJET Then
 Call rstProjSoum.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)
1  Else
 Call rstProjSoum.Open("SELECT * FROM GrbSoumissionElec WHERE IDSoumission = '" & sNoProjSoum & "'", g_connData, adOpenDynamic, adLockOptimistic)
 End If

If Not rstProjSoum.EOF Then
 If m_eType = TYPE_PROJET Then
 If Right$(sNoProjSoum, 2) = "99" Then
 sFilterNoProjet = "LEFT(NoProjet, 6) = '" & Left$(sNoProjSoum, 6) & "'"
 Else
 sFilterNoProjet = "NoProjet = '" & sNoProjSoum & "'"
 End If

 sDateDebut = "TIMESERIAL(Left(GrbPunch.HeureDébut,2),RIGHT(GrbPunch.HeureDébut,2),0)"

 sDateFin = "TIMESERIAL(Left(GrbPunch.HeureFin,2),RIGHT(GrbPunch.HeureFin,2),0)"

 sTotal = "(SUM(" & sDateFin & " - " & sDateDebut & ")* 24) As Total"

 Set rstPunch = New ADODB.Recordset

 Call rstPunch.Open("SELECT Type, " & sTotal & " FROM GrbPunch WHERE " & sFilterNoProjet & " And HeureFin is not null AND HeureDébut is not null GROUP BY Type", g_connData, adOpenDynamic, adLockOptimistic)

 dblTotalDessin = 0
 dblTotalFabrication = 0
 dblTotalAssemblage = 0
 dblTotalProgInterface = 0
 dblTotalProgAutomate = 0
 dblTotalProgRobot = 0
 dblTotalVision = 0
dblTotalTest = 0
 dblTotalInstallation = 0
 dblTotalMiseService = 0
 dblTotalFormation = 0
 dblTotalGestion = 0
 dblTotalShipping = 0

 Do While Not rstPunch.EOF
 If Not IsNull(rstPunch.Fields("Total")) Then
 Select Case rstPunch.Fields("Type")
 Case "Dessin":
 If Not IsNull(rstProjSoum.Fields("TauxDessin")) Then
 dblTotalDessin = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxDessin"))
 Else
 dblTotalDessin = 0
 End If

 Case "Fabrication":
 If Not IsNull(rstProjSoum.Fields("TauxFabrication")) Then
 dblTotalFabrication = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxFabrication"))
 Else
 dblTotalFabrication = 0
 End If
 
4 Case "Assemblage":
4 If Not IsNull(rstProjSoum.Fields("TauxAssemblage")) Then
4 dblTotalAssemblage = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxAssemblage"))
4 Else
4 dblTotalAssemblage = 0
4 End If
 
4 Case "ProgInterface":
4 If Not IsNull(rstProjSoum.Fields("TauxProgInterface")) Then
4 dblTotalProgInterface = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxProgInterface"))
4 Else
4 dblTotalProgInterface = 0
4  End If
 
4  Case "ProgAutomate":
4  If Not IsNull(rstProjSoum.Fields("TauxProgAutomate")) Then
4  dblTotalProgAutomate = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxProgAutomate"))
4  Else
4  dblTotalProgAutomate = 0
4  End If
 
4  Case "ProgRobot":
50 If Not IsNull(rstProjSoum.Fields("TauxProgRobot")) Then
 dblTotalProgRobot = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxProgRobot"))
 Else
 dblTotalProgRobot = 0
 End If
 
 Case "Vision":
 If Not IsNull(rstProjSoum.Fields("TauxVision")) Then
 dblTotalVision = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxVision"))
 Else
 dblTotalVision = 0
 End If
 
 Case "Test":
5  If Not IsNull(rstProjSoum.Fields("TauxTest")) Then
5  dblTotalTest = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxTest"))
5  Else
5  dblTotalTest = 0
5  End If
 
5  Case "Installation":
5  If Not IsNull(rstProjSoum.Fields("TauxInstallation")) Then
5  dblTotalInstallation = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxInstallation"))
60 Else
  dblTotalInstallation = 0
  End If
 
  Case "MiseService":
  If Not IsNull(rstProjSoum.Fields("TauxMiseService")) Then
  dblTotalMiseService = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxMiseService"))
  Else
  dblTotalMiseService = 0
  End If
 
  Case "Formation":
  If Not IsNull(rstProjSoum.Fields("TauxFormation")) Then
  dblTotalFormation = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxFormation"))
6  Else
6  dblTotalFormation = 0
6  End If
 
6  Case "Gestion":
6  If Not IsNull(rstProjSoum.Fields("TauxGestion")) Then
6  dblTotalGestion = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxGestion"))
6  Else
6  dblTotalGestion = 0
70 End If
 
  Case "Shipping":
  If Not IsNull(rstProjSoum.Fields("TauxShipping")) Then
  dblTotalShipping = CDbl(rstPunch.Fields("Total")) * CDbl(rstProjSoum.Fields("TauxShipping"))
  Else
  dblTotalShipping = 0
  End If
  End Select
  End If
 
  Call rstPunch.MoveNext
  Loop

  Call rstPunch.Close
   Set rstPunch = Nothing
   Else
7  If Not IsNull(rstProjSoum.Fields("TempsDessin")) And Not IsNull(rstProjSoum.Fields("TauxDessin")) Then
7  dblTotalDessin = CDbl(rstProjSoum.Fields("TempsDessin")) * CDbl(rstProjSoum.Fields("TauxDessin"))
7  Else
7  dblTotalDessin = 0
7  End If

7  If rstProjSoum.Fields("SansTemps") = False Then
80 If Not IsNull(rstProjSoum.Fields("TempsFabrication")) And Not IsNull(rstProjSoum.Fields("TauxFabrication")) Then
  dblTotalFabrication = CDbl(rstProjSoum.Fields("TempsFabrication")) * CDbl(rstProjSoum.Fields("TauxFabrication"))
  Else
  dblTotalFabrication = 0
  End If
  Else
  dblTotalFabrication = 0
  End If

  If Not IsNull(rstProjSoum.Fields("TempsAssemblage")) And Not IsNull(rstProjSoum.Fields("TauxAssemblage")) Then
  dblTotalAssemblage = CDbl(rstProjSoum.Fields("TempsAssemblage")) * CDbl(rstProjSoum.Fields("TauxAssemblage"))
  Else
  dblTotalAssemblage = 0
   End If

   If Not IsNull(rstProjSoum.Fields("TempsProgInterface")) And Not IsNull(rstProjSoum.Fields("TauxProgInterface")) Then
   dblTotalProgInterface = CDbl(rstProjSoum.Fields("TempsProgInterface")) * CDbl(rstProjSoum.Fields("TauxProgInterface"))
   Else
8  dblTotalProgInterface = 0
8  End If

8  If Not IsNull(rstProjSoum.Fields("TempsProgAutomate")) And Not IsNull(rstProjSoum.Fields("TauxProgAutomate")) Then
8  dblTotalProgAutomate = CDbl(rstProjSoum.Fields("TempsProgAutomate")) * CDbl(rstProjSoum.Fields("TauxProgAutomate"))
90 Else
  dblTotalProgAutomate = 0
  End If

  If Not IsNull(rstProjSoum.Fields("TempsProgRobot")) And Not IsNull(rstProjSoum.Fields("TauxProgRobot")) Then
  dblTotalProgRobot = CDbl(rstProjSoum.Fields("TempsProgRobot")) * CDbl(rstProjSoum.Fields("TauxProgRobot"))
  Else
  dblTotalProgRobot = 0
  End If

  If Not IsNull(rstProjSoum.Fields("TempsVision")) And Not IsNull(rstProjSoum.Fields("TauxVision")) Then
  dblTotalVision = CDbl(rstProjSoum.Fields("TempsVision")) * CDbl(rstProjSoum.Fields("TauxVision"))
  Else
  dblTotalVision = 0
 End If

   If Not IsNull(rstProjSoum.Fields("TempsTest")) And Not IsNull(rstProjSoum.Fields("TauxTest")) Then
 dblTotalTest = CDbl(rstProjSoum.Fields("TempsTest")) * CDbl(rstProjSoum.Fields("TauxTest"))
   Else
 dblTotalTest = 0
   End If

 If Not IsNull(rstProjSoum.Fields("TempsInstallation")) And Not IsNull(rstProjSoum.Fields("TauxInstallation")) Then
9  dblTotalInstallation = CDbl(rstProjSoum.Fields("TempsInstallation")) * CDbl(rstProjSoum.Fields("TauxInstallation"))
 Else
 dblTotalInstallation = 0
1 End If

1 If Not IsNull(rstProjSoum.Fields("TempsMiseService")) And Not IsNull(rstProjSoum.Fields("TauxMiseService")) Then
 dblTotalMiseService = CDbl(rstProjSoum.Fields("TempsMiseService")) * CDbl(rstProjSoum.Fields("TauxMiseService"))
1 Else
 dblTotalMiseService = 0
1 End If

 If Not IsNull(rstProjSoum.Fields("TempsFormation")) And Not IsNull(rstProjSoum.Fields("TauxFormation")) Then
 dblTotalFormation = CDbl(rstProjSoum.Fields("TempsFormation")) * CDbl(rstProjSoum.Fields("TauxFormation"))
 Else
 dblTotalFormation = 0
10  End If
 
10  If Not IsNull(rstProjSoum.Fields("TempsGestion")) And Not IsNull(rstProjSoum.Fields("TauxGestion")) Then
10  dblTotalGestion = CDbl(rstProjSoum.Fields("TempsGestion")) * CDbl(rstProjSoum.Fields("TauxGestion"))
10  Else
10  dblTotalGestion = 0
10  End If

10  If Not IsNull(rstProjSoum.Fields("TempsShipping")) And Not IsNull(rstProjSoum.Fields("TauxShipping")) Then
10  dblTotalShipping = CDbl(rstProjSoum.Fields("TempsShipping")) * CDbl(rstProjSoum.Fields("TauxShipping"))
1 Else
1 dblTotalShipping = 0
1 End If
1 End If

1 dblTotalTemps = dblTotalDessin + _
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

1 Set rstPiece = New ADODB.Recordset

1 If m_eType = TYPE_PROJET Then
1 Call rstPiece.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & sNoProjSoum & "' AND Type = 'E'", g_connData, adOpenDynamic, adLockOptimistic)
1 Else
1 Call rstPiece.Open("SELECT * FROM GrbSoumission_Pieces WHERE IDSoumission = '" & sNoProjSoum & "' AND Type = 'E'", g_connData, adOpenDynamic, adLockOptimistic)
1 End If

 'Pour chaque élément du recordset
1 Do While Not rstPiece.EOF
1 If Trim(rstPiece.Fields("Prix_total")) <> vbNullString Then
 'On additionne le prix total
1 dblPrixPieces = dblPrixPieces + CDbl(rstPiece.Fields("Prix_total")) - CDbl(rstPiece.Fields("Profit_Argent"))
 
 'On additionne le profit
 dblProfit = dblProfit + CDbl(rstPiece.Fields("Profit_Argent"))
1 End If

 Call rstPiece.MoveNext
1 Loop

 Call rstPiece.Close
11  Set rstPiece = Nothing

 If m_eType = TYPE_PROJET Then
 dblHebergement = 0
1 dblRepas = 0
1 dblTransport = 0
1 dblUniteMobile = 0
1 Else
1 If Not IsNull(rstProjSoum.Fields("TotalHebergement")) Then
1 dblHebergement = CDbl(rstProjSoum.Fields("TotalHebergement"))
1 Else
1 dblHebergement = 0
1 End If

1 If Not IsNull(rstProjSoum.Fields("TotalRepas")) Then
1 dblRepas = CDbl(rstProjSoum.Fields("TotalRepas"))
1 Else
1 dblRepas = 0
1 End If
 
1 If Not IsNull(rstProjSoum.Fields("TempsTransport")) Then
1 dblTransport = CDbl(rstProjSoum.Fields("TempsTransport")) * CDbl(rstProjSoum.Fields("TauxTransport"))
1 Else
1 dblTransport = 0
1 End If

13 If Not IsNull(rstProjSoum.Fields("TempsUniteMobile")) Then
1 dblUniteMobile = CDbl(rstProjSoum.Fields("TempsUniteMobile")) * CDbl(rstProjSoum.Fields("TauxUniteMobile"))
1 Else
1 dblUniteMobile = 0
1 End If
1 End If

1 If IsNumeric(rstProjSoum.Fields("PrixEmballage")) Then
1 dblPrixEmballage = CDbl(rstProjSoum.Fields("PrixEmballage"))
1 Else
1 dblPrixEmballage = 0
1 End If
 
13  dblTotalResteTemps = dblHebergement + dblRepas + dblTransport + dblUniteMobile + dblPrixEmballage

1 If IsNumeric(rstProjSoum.Fields("total_manuel")) Then
1 dblTotalManuel = CDbl(rstProjSoum.Fields("total_manuel"))
1 Else
1 dblTotalManuel = 0
1 End If

1dblTotalPieceImprevue = (dblPrixPieces + dblProfit) * (1 + CDbl(rstProjSoum.Fields("Imprevue")))

1 dblPrixTotal = dblTotalTemps + dblTotalManuel + dblTotalPieceImprevue + dblTotalResteTemps

 'Le calcul de la commission est sur les manuels (Nbre de page * prix par pages * nbre de copies) + (prix des pièces * pourcentage d'imprévus) + total des temps
140 dblCommission = dblPrixTotal * CDbl(rstProjSoum.Fields("Commission"))

 'Le prix total est le calcul des manuels (Nbre de page * prix par pages * nbre de copies) + (prix des pièces * pourcentage d'imprévus) + total des temps + total de la commission
14dblGrandTotal = dblPrixTotal + dblCommission

 'Format monétaire avec 2 chiffres après la virgule
14 rstProjSoum.Fields("total_commission") = dblCommission
14 rstProjSoum.Fields("Total_manuel") = dblTotalManuel
14 rstProjSoum.Fields("Total_temps") = dblTotalTemps
14 rstProjSoum.Fields("total_imprevue") = dblTotalPieceImprevue - (dblPrixPieces + dblProfit)
14 rstProjSoum.Fields("total_piece") = dblPrixPieces
14 rstProjSoum.Fields("Total_Commission") = Conversion(CStr(Round(dblCommission, 2)), MODE_ARGENT)
14 rstProjSoum.Fields("PrixTotal") = Conversion(CStr(Round(dblGrandTotal, 2)), MODE_ARGENT)
14 rstProjSoum.Fields("Total_Profit") = Conversion(CStr(Round(dblProfit, 2)), MODE_ARGENT)

14 Call rstProjSoum.Update
14 Else
14  If m_eType = TYPE_PROJET Then
14  Call MsgBox("Le projet " & sNoProjSoum & " est inexistant!", vbOKOnly, "Erreur")
14  Else
14  Call MsgBox("La soumission " & sNoProjSoum & " est inexistant!", vbOKOnly, "Erreur")
14  End If
14  End If

14  Call rstProjSoum.Close
14  Set rstProjSoum = Nothing

150Exit Sub

Oups:

150 woups"frmProjSoumElec", "CalculerTotalRecordset", Err, Erl, sNoProjSoum)
End Sub

Private Sub CalculerPrixFacturation(ByVal sNoFacturation As String, ByRef sCommission As String, ByRef sPrixTotal As String, ByRef sProfit As String, ByRef sTempsFabrication As String, ByRef sTotalPiece As String, ByRef sImprevue As String, ByRef sTotalTemps As String, ByRef sManuel As String)

 On Error GoTo Oups

 'Méthode pour calculer le prix
 Dim iCompteur As Integer
 Dim dblTotalDessin As Double
 Dim dblTotalFabrication As Double
 Dim dblTotalAssemblage As Double
 Dim dblTotalProgInterface As Double
 Dim dblTotalProgAutomate As Double
 Dim dblTotalProgRobot As Double
 Dim dblTotalVision As Double
 Dim dblTotalTest As Double
 Dim dblTotalInstallation As Double
  Dim dblTotalMiseService As Double
  Dim dblTotalFormation As Double
  Dim dblTotalGestion As Double
  Dim dblTotalShipping As Double
  Dim dblPrixPieces As Double
  Dim dblPrixTotal As Double
  Dim dblCommission As Double
  Dim dblTotalTemps As Double
10 Dim dblProfit As Double
Dim dblTotalManuel As Double
Dim dblTotalPieceImprevue As Double
Dim dblGrandTotal As Double
Dim dblTempsFabrication As Double
 
dblTotalDessin = CDbl(m_sTempsDessin) * CDbl(m_sTauxDessin)

If m_bSansTemps = False Then
 dblTotalFabrication = CDbl(m_sTempsFabrication) * CDbl(m_sTauxFabrication)
Else
 dblTotalFabrication = 0
End If

dblTotalAssemblage = CDbl(m_sTempsAssemblage) * CDbl(m_sTauxAssemblage)
1  dblTotalProgInterface = CDbl(m_sTempsProgInterface) * CDbl(m_sTauxProgInterface)
dblTotalProgAutomate = CDbl(m_sTempsProgAutomate) * CDbl(m_sTauxProgAutomate)
 dblTotalProgRobot = CDbl(m_sTempsProgRobot) * CDbl(m_sTauxProgRobot)
dblTotalVision = CDbl(m_sTempsVision) * CDbl(m_sTauxVision)
 dblTotalTest = CDbl(m_sTempsTest) * CDbl(m_sTauxTest)
dblTotalInstallation = CDbl(m_sTempsInstallation) * CDbl(m_sTauxInstallation)
 dblTotalMiseService = CDbl(m_sTempsMiseService) * CDbl(m_sTauxMiseService)
1  dblTotalFormation = CDbl(m_sTempsFormation) * CDbl(m_sTauxFormation)
 dblTotalGestion = CDbl(m_sTempsGestion) * CDbl(m_sTauxGestion)
 dblTotalShipping = CDbl(m_sTempsShipping) * CDbl(m_sTauxShipping)
 
dblTotalTemps = dblTotalDessin + _
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
For iCompteur = 1 To lvwSoumission.ListItems.count
 'Si ce n'est pas une section
 If lvwSoumission.ListItems(iCompteur).Tag <> vbNullString Then
 'Si ce n'est pas une sous-section
 If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> vbNullString And lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> "Texte" And lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> "Text" Then
 If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_FACTURATION) = sNoFacturation Then
 If Trim$(lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_TOTAL)) <> vbNullString Then
 'On additionne le prix total
 dblPrixPieces = dblPrixPieces + lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_TOTAL)
 
 'On additionne le profit
 dblProfit = dblProfit + lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PROFIT)

 'Calcul des heures de fabrication
 If m_bSansTemps = False Then
 If Trim$(lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_MONTAGE)) <> vbNullString Then
 dblTempsFabrication = dblTempsFabrication + lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_MONTAGE)
 End If
 End If
 End If
 End If
 End If
End If
Next
 
30 If IsNumeric(txtPrixManuel.Text) Then
3 dblTotalManuel = CDbl(txtPrixManuel.Text)
Else
 dblTotalManuel = 0
End If
 
dblTotalPieceImprevue = dblPrixPieces * (1 + CDbl(m_sImprevue))
 
dblPrixTotal = dblTotalTemps + dblTotalManuel + dblTotalPieceImprevue
 
 'Le calcul de la commission est sur les manuels (Nbre de page * prix par pages * nbre de copies) + (prix des pièces * pourcentage d'imprévus) + total des temps
dblCommission = dblPrixTotal * CDbl(m_sCommission)
 
 'Le prix total est le calcul des manuels (Nbre de page * prix par pages * nbre de copies) + (prix des pièces * pourcentage d'imprévus) + total des temps + total de la commission
dblGrandTotal = dblPrixTotal + dblCommission
 
 'Format monétaires avec 2 chiffres après la virgule
sCommission = Conversion(CStr(Round(dblCommission, 2)), MODE_ARGENT)
sPrixTotal = Conversion(CStr(Round(dblGrandTotal, 2)), MODE_ARGENT)
sProfit = Conversion(CStr(Round(dblProfit, 2)), MODE_ARGENT)
3  sTempsFabrication = dblTempsFabrication
sImprevue = Conversion(CStr(Round(dblPrixPieces * CDbl(m_sImprevue), 2)), MODE_ARGENT)
3  sManuel = Conversion(CStr(Round(dblTotalManuel, 2)), MODE_ARGENT)
sTotalPiece = Conversion(CStr(Round(dblPrixPieces, 2)), MODE_ARGENT)
3  sTotalTemps = Conversion(CStr(Round(dblTotalTemps, 2)), MODE_ARGENT)

Exit Sub

Oups:

3  wOups "frmProjSoumElec", "CalculerPrix", Err, Err.number, Err.Description
End Sub

Private Sub ChoisirFournisseur()

 On Error GoTo Oups

 'On ajoute la pièce dans lvwSoumission
 Dim sQuantite As String
 Dim sSousSection As String
 Dim bDemanderSS As Boolean
 Dim sParams As String
 
 'Si l'utilisateur a déjà choisi un emplacement, il ne faut pas
 'lui demander dans quelle sous-section
 
 'Si il y a des enregistrements dans le listview
 If lvwSoumission.ListItems.count > 0 Then
 'Si le premier n'est pas sélectionné.. celui-ci est sélectionné par défaut
 If lvwSoumission.SelectedItem.Index > 1 Then
 'Si l'emplacement est valide
 If VerifierEmplacement(lvwSoumission.SelectedItem.Index) = True Then
 'Si c'est une sous-section
 If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).Tag = vbNullString Then
 'Si l'autre d'au dessus est une section
 If lvwSoumission.ListItems(lvwSoumission.SelectedItem.Index - 1).Tag = vbNullString Then
 'Message d'erreur
 Call MsgBox("Vous ne pouvez pas mettre une pièce entre une section et une sous-section", vbOKOnly, "Erreur")
 
  frafournisseur.Visible = False
 
 'Il faut resélectionné le premier pour faire comme s'il n'était plus
 'sélectionné
  Call Deselect
 
  lvwSoumission.ListItems(1).Selected = True
 
  Exit Sub
  Else
 'Sinon, on prend le tag de la section d'en haut
  sSousSection = lvwSoumission.ListItems(lvwSoumission.SelectedItem.Index - 1).ListSubItems(I_COL_SOUM_PIECE).Tag
  End If
  Else
 'On prend le tag de l'élément sélectionné
 sSousSection = lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).Tag
 End If
 Else
 If lvwSoumission.ListItems(lvwSoumission.SelectedItem.Index - 1).ListSubItems(I_COL_SOUM_MANUFACT).Tag <> "" Then
 If MsgBox("Vous essayez d'ajouter une pièce de la section " & cmbSections.Text & " dans la section " & cmbSections.LIST(lvwSoumission.ListItems(lvwSoumission.SelectedItem.Index - 1).ListSubItems(I_COL_SOUM_MANUFACT).Tag - 1) & vbNewLine & "Voulez-vous ajouter la pièce dans la section " & cmbSections.LIST(lvwSoumission.ListItems(lvwSoumission.SelectedItem.Index - 1).ListSubItems(I_COL_SOUM_MANUFACT).Tag - 1), vbYesNo, "Erreur") = vbYes Then
 cmbSections.ListIndex = lvwSoumission.ListItems(lvwSoumission.SelectedItem.Index - 1).ListSubItems(I_COL_SOUM_MANUFACT).Tag - 1

 Call ChoisirFournisseur
 End If
 
 frafournisseur.Visible = False
 
 'Il faut resélectionné le premier pour faire comme si il n'était plus
 'sélectionné
 Call Deselect
 
 lvwSoumission.ListItems(1).Selected = True
 
 Exit Sub
 Else
 Call MsgBox("Impossible d'ajouter entre une section et une sous-section!", vbOKOnly, "Erreur")

 Exit Sub
 End If
 End If
 Else
 bDemanderSS = True
1  End If
 Else
 'Sinon, on demande la section
 bDemanderSS = True
End If
 
 'Saisie de la quantité
sQuantite = InputBox("Quelle est la quantité?")

sQuantite = Replace(sQuantite, ".", ",")
 
If sQuantite <> vbNullString Then
 If Not IsNumeric(sQuantite) Then
 Call MsgBox("Quantité non numérique!", vbOKOnly, "Erreur")
 
 Exit Sub
 Else
 If sQuantite < 0 Then
 If lvwfournisseur.SelectedItem.Text = "CHOISIR ULTÉRIEUREMENT" Then
 Call MsgBox("Impossible de faire une demande de prix sur une pièce négative!", vbOKOnly, "Erreur")

 Exit Sub
 End If
 End If
End If
Else
Exit Sub
End If

30 If bDemanderSS = True Then
3 If m_sSousSection <> S_PAS_SOUS_SECTION Then
 sSousSection = InputBox("Quelle est la sous-section?", , m_sSousSection)
 Else
 sSousSection = InputBox("Quelle est la sous-section?")
 End If
End If
 
 'Si la sous-section est vide
If sSousSection = vbNullString Then
 'On initialise la sous-section à "PAS DE SOUS-SECTIONS"
 sSousSection = S_PAS_SOUS_SECTION
 m_sSousSection = vbNullString
Else
 m_sSousSection = sSousSection
3  End If

If sQuantite < 0 Then
If m_eType = TYPE_PROJET Then
 If Right$(txtNoProjSoum.Text, 2) >= 60 And Right$(txtNoProjSoum.Text, 2) <=   Then
 Call AjouterNegatifDansListView(CDbl(sQuantite), sSousSection)
 Else
 Call AjouterDansListViewSoumission(CDbl(sQuantite), sSousSection)
 End If
Else
4 Call AjouterDansListViewSoumission(CDbl(sQuantite), sSousSection)
4 End If
4 Else
4 Call AjouterDansListViewSoumission(CDbl(sQuantite), sSousSection)
4 End If
 
 'Calcul des prix
4 Call CalculerPrix
 
 'On cache le listview
4 frafournisseur.Visible = False
 
 'Resélectionne le premier élément du listview
4 If lvwSoumission.ListItems.count > 0 Then
4 Call Deselect

4 lvwSoumission.ListItems(1).Selected = True
4 End If

4  Exit Sub

Oups:

4  If Err.number = 13 And Erl = 110 Then
4  sParams = "cmbSections.Text : " & cmbSections.Text & " " & _
 "No Proj/Soum : " & txtNoProjSoum.Text & " " & _
 "lvwSoumission.SelectedItem.Index - 1 : " & lvwSoumission.SelectedItem.Index - 1 & " " & _
 "lvwSoumission.ListItems(lvwSoumission.SelectedItem.Index - 1).ListSubItems(I_COL_SOUM_MANUFACT).Tag : " & lvwSoumission.ListItems(lvwSoumission.SelectedItem.Index - 1).ListSubItems(I_COL_SOUM_MANUFACT).Tag
 
4  woups"frmProjSoumElec", "ChoisirFournisseur", Err, Erl, sParams)
4  Else
4  wOups "frmProjSoumElec", "ChoisirFournisseur", Err, Err.number, Err.Description
4  End If
End Sub

Private Sub ChoisirFournisseurMateriel()

 On Error GoTo Oups

 'On ajoute la pièce en négatif dans le ListView
 Dim rstProjet As ADODB.Recordset
 Dim rstConfig As ADODB.Recordset
 Dim itmAncien As ListItem
 Dim itmNouveau As ListItem
 Dim sQuantite As String
 Dim sExtra As String
 Dim sTauxUSA As String
 Dim sTauxSPA As String

 If m_bChangementFRS = True Then
 If lvwfournisseur.SelectedItem.Text <> "CHOISIR ULTÉRIEUREMENT" Then
  Set rstConfig = New ADODB.Recordset

  Call rstConfig.Open("SELECT TauxAmericain, TauxEspagnol FROM GrbConfig", g_connData, adOpenForwardOnly, adLockReadOnly)

  sTauxUSA = rstConfig.Fields("TauxAmericain")
  sTauxSPA = rstConfig.Fields("TauxEspagnol")

  Call rstConfig.Close
  Set rstConfig = Nothing

  lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_DISTRIB) = lvwfournisseur.SelectedItem.Text
  lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_DISTRIB).Tag = lvwfournisseur.SelectedItem.Tag

 If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor <> COLOR_BRUN Then
 'Prix listé
 If Trim$(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_LIST)) = vbNullString Then
 lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion("0", MODE_ARGENT, 4)
 Else
 If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "USA" Then
 lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_LIST)) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
 Else
 If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "SPA" Then
 lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_LIST)) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
 Else
 lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_LIST), MODE_ARGENT, 4)
 End If
 End If
 End If

 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PRIX_LIST).Tag
 
 'Si il y a un prix net, on le met l'escompte et le prix net sinon, on prend le prix
 'spécial pour mettre dans le prix net
 If Trim$(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_NET)) <> vbNullString Then
 lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_ESCOMPTE), MODE_POURCENT)

 If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "USA" Then
 lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_NET)) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
1  Else
 If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "SPA" Then
 lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_NET)) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
 Else
 lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_NET), MODE_ARGENT, 4)
 End If
 End If
 Else
 If Trim$(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_SP)) <> vbNullString Then
 If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "USA" Then
 lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_SP)) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
 Else
 If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "SPA" Then
 lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_SP)) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
 Else
 lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_SP), MODE_ARGENT, 4)
 End If
 End If
 Else
 lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion("0", MODE_POURCENT)
 lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PRIX_NET) = Conversion("0", MODE_ARGENT, 4)
 End If
 End If

 lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_TOTAL) = Conversion(Round(Replace(lvwSoumission.SelectedItem.Text, "*", vbNullString) * lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PRIX_NET) * CSng(m_sProfit), 2), MODE_ARGENT, 2)

 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_TOTAL).Tag = lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag
 
 'Pour le profit, c'est le prix total - (prix net * quantité)
 lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PROFIT) = Conversion(CStr(Round(lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_TOTAL) - (lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PRIX_NET) * Replace(lvwSoumission.SelectedItem.Text, "*", vbNullString)), 2)), MODE_ARGENT)
 End If
 Else
 lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_DISTRIB) = vbNullString
 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_DISTRIB).Tag = 0

 lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion("0", MODE_ARGENT, 2)
 lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion("0", MODE_POURCENT)
 lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PRIX_NET) = Conversion("0", MODE_ARGENT, 2)

 lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_TOTAL) = Conversion(Round(Replace(lvwSoumission.SelectedItem.Text, "*", vbNullString) * lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PRIX_NET) * CSng(m_sProfit), 2), MODE_ARGENT, 2)
 
 lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PROFIT) = Conversion(CStr(Round(lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_TOTAL) - (lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PRIX_NET) * Replace(lvwSoumission.SelectedItem.Text, "*", vbNullString)), 2)), MODE_ARGENT)

 If m_eType = TYPE_PROJET Then
 If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_COMMENTAIRE) <> "" Then
 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_COMMENTAIRE).ForeColor = COLOR_MAGENTA
 End If

 If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_FACTURATION) <> "" Then
 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_FACTURATION).ForeColor = COLOR_MAGENTA
 End If

4 If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_ID) <> "" Then
4 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_ID).ForeColor = COLOR_MAGENTA
4 End If
4 End If

4 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_DESCR).ForeColor = COLOR_MAGENTA
4 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = COLOR_MAGENTA
4 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = COLOR_MAGENTA
4 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_MONTAGE).ForeColor = COLOR_MAGENTA
4 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_MAGENTA
4 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = COLOR_MAGENTA
4 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = COLOR_MAGENTA
4  lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = COLOR_MAGENTA
4  lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_TEMPS).ForeColor = COLOR_MAGENTA
4  lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = COLOR_MAGENTA

4  Call lvwSoumission.Refresh
4  End If

4  Call CalculerPrix
 
 'On cache le listview
4  frafournisseur.Visible = False

4  m_bPieceInutile = False
50 m_bChangementFRS = False
50 Else
 If Right$(txtNoProjSoum.Text, 2) >= "00" And Right$(txtNoProjSoum.Text, 2) <= "19" Then
 sExtra = InputBox("Dans quel extra le retour doit être fait ? (2 chiffres seulement)")

 If Len(sExtra) <> 2 Then
 Call MsgBox("Format incorrect!", vbOKOnly, "Erreur")

 Exit Sub
 End If

 If Not IsNumeric(sExtra) Then
 Call MsgBox("L'extra doit être numérique!", vbOKOnly, "Erreur")

 Exit Sub
 End If

5  If sExtra < 60 Or sExtra >   Then
5  Call MsgBox("L'extra doit être entre 60 et 98!", vbOKOnly, "Erreur")

5  Exit Sub
5  End If

5  Set rstProjet = New ADODB.Recordset

5  Call rstProjet.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sExtra & "'", g_connData, adOpenDynamic, adLockOptimistic)

5  If rstProjet.EOF Then
5  Call MsgBox("Le projet " & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sExtra & " n'existe pas!", vbOKOnly, "Erreur")

60 Call rstProjet.Close
  Set rstProjet = Nothing

  Exit Sub
  Else
  Call rstProjet.Close
  Set rstProjet = Nothing
  End If
  End If

 'Saisie de la quantité
  sQuantite = InputBox("Quelle est la quantité?")

  sQuantite = Replace(sQuantite, ".", ",")

  sQuantite = Replace(sQuantite, "-", "")

  If sQuantite <> vbNullString Then
6  If Not IsNumeric(sQuantite) Then
6  Call MsgBox("Quantité non numérique!", vbOKOnly, "Erreur")

6  Exit Sub
6  End If
6  Else
6  Exit Sub
6  End If

6  If CDbl(sQuantite) <= CDbl(Replace(lvwSoumission.SelectedItem.Text, "*", vbNullString)) Then
70 Set itmAncien = lvwSoumission.SelectedItem
  Set itmNouveau = lvwSoumission.ListItems.Add(itmAncien.Index + 1)

  itmNouveau.Checked = itmAncien.Checked

 'Quantité
  itmNouveau.Text = "-" & sQuantite

 'On met l'id de la section dans le tag du listItem
  itmNouveau.Tag = itmAncien.Tag

 'No d'item
  itmNouveau.SubItems(I_COL_SOUM_PIECE) = itmAncien.SubItems(I_COL_SOUM_PIECE)

 'On met le nom de la sous-section dans le tag du no d'item
  itmNouveau.ListSubItems(I_COL_SOUM_PIECE).Tag = itmAncien.ListSubItems(I_COL_SOUM_PIECE).Tag

 'On met la description en francais dans la colonne et la description en anglais
 'dans le tag
  itmNouveau.SubItems(I_COL_SOUM_DESCR) = itmAncien.SubItems(I_COL_SOUM_DESCR)
  itmNouveau.ListSubItems(I_COL_SOUM_DESCR).Tag = itmAncien.ListSubItems(I_COL_SOUM_DESCR).Tag

 'On met le nom du fabricant dans la colonne et l'ordre de la section dans le tag
  itmNouveau.SubItems(I_COL_SOUM_MANUFACT) = itmAncien.SubItems(I_COL_SOUM_MANUFACT)
  itmNouveau.ListSubItems(I_COL_SOUM_MANUFACT).Tag = itmAncien.ListSubItems(I_COL_SOUM_MANUFACT).Tag
 
 'Prix listé
  itmNouveau.SubItems(I_COL_SOUM_PRIX_LIST) = itmAncien.SubItems(I_COL_SOUM_PRIX_LIST)

   itmNouveau.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = itmAncien.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag

   itmNouveau.SubItems(I_COL_SOUM_ESCOMPTE) = itmAncien.SubItems(I_COL_SOUM_ESCOMPTE)

7  itmNouveau.SubItems(I_COL_SOUM_PRIX_NET) = itmAncien.SubItems(I_COL_SOUM_PRIX_NET)

 'On met le fournisseur dans la colonne et l'id dans le tag
7  itmNouveau.SubItems(I_COL_SOUM_DISTRIB) = lvwfournisseur.SelectedItem.Text
7  itmNouveau.ListSubItems(I_COL_SOUM_DISTRIB).Tag = lvwfournisseur.SelectedItem.Tag

 'Temps
7  itmNouveau.SubItems(I_COL_SOUM_TEMPS) = itmAncien.SubItems(I_COL_SOUM_TEMPS)

 'Si le temps n'est pas vide
7  If Trim$(itmNouveau.SubItems(I_COL_SOUM_TEMPS)) <> vbNullString Then
 'On calcul le temps * quantité pour la colonne montage
7  itmNouveau.SubItems(I_COL_SOUM_MONTAGE) = CDbl(Replace(itmNouveau.SubItems(I_COL_SOUM_TEMPS), ".", ",")) * CDbl(Replace(itmNouveau.Text, "*", vbNullString))
80 Else
  itmNouveau.SubItems(I_COL_SOUM_MONTAGE) = vbNullString
  End If

 'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
  itmNouveau.SubItems(I_COL_SOUM_TOTAL) = Conversion(Round(CDbl(Replace(itmNouveau.Text, "*", "")) * CDbl(itmNouveau.SubItems(I_COL_SOUM_PRIX_NET)) * CDbl(m_sProfit), 2), MODE_ARGENT)

 'Pour le profit, c'est le prix total - (prix net * quantité)
  itmNouveau.SubItems(I_COL_SOUM_PROFIT) = Conversion(Round(CDbl(itmNouveau.SubItems(I_COL_SOUM_TOTAL)) - (CDbl(Replace(itmNouveau.Text, "*", "") * CDbl(itmNouveau.SubItems(I_COL_SOUM_PRIX_NET)))), 2), MODE_ARGENT)

  If Right$(txtNoProjSoum.Text, 2) >= "00" And Right$(txtNoProjSoum.Text, 2) <= "19" Then
 'Pour savoir lors de l'enregistrement qu'il faut le lier avec un extra
  itmNouveau.ListSubItems(I_COL_SOUM_PROFIT).Tag = "RETOUR " & sExtra
  End If

  itmNouveau.SubItems(I_COL_SOUM_DATE_COMMANDE) = " "
  itmNouveau.SubItems(I_COL_SOUM_DATE_REQUISE) = " "
  itmNouveau.SubItems(I_COL_SOUM_NOM_COMMANDE) = " "
  itmNouveau.SubItems(I_COL_SOUM_NO_SEQUENTIEL) = " "

   If itmAncien.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_NOIR Then
   itmNouveau.ListSubItems(I_COL_SOUM_DATE_COMMANDE).ForeColor = COLOR_NOIR
   itmNouveau.ListSubItems(I_COL_SOUM_DATE_REQUISE).ForeColor = COLOR_NOIR
   itmNouveau.ListSubItems(I_COL_SOUM_DESCR).ForeColor = COLOR_NOIR
8  itmNouveau.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = COLOR_NOIR
8  itmNouveau.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = COLOR_NOIR
8  itmNouveau.ListSubItems(I_COL_SOUM_ID).ForeColor = COLOR_NOIR
8  itmNouveau.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = COLOR_NOIR
90 itmNouveau.ListSubItems(I_COL_SOUM_MONTAGE).ForeColor = COLOR_NOIR
  itmNouveau.ListSubItems(I_COL_SOUM_NO_SEQUENTIEL).ForeColor = COLOR_NOIR
  itmNouveau.ListSubItems(I_COL_SOUM_NOM_COMMANDE).ForeColor = COLOR_NOIR
  itmNouveau.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_NOIR
  itmNouveau.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = COLOR_NOIR
  itmNouveau.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = COLOR_NOIR
  itmNouveau.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = COLOR_NOIR
  itmNouveau.ListSubItems(I_COL_SOUM_TEMPS).ForeColor = COLOR_NOIR
  itmNouveau.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = COLOR_NOIR
  Else
  itmNouveau.ListSubItems(I_COL_SOUM_DATE_COMMANDE).ForeColor = COLOR_BRUN
  itmNouveau.ListSubItems(I_COL_SOUM_DATE_REQUISE).ForeColor = COLOR_BRUN
 itmNouveau.ListSubItems(I_COL_SOUM_DESCR).ForeColor = COLOR_BRUN
   itmNouveau.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = COLOR_BRUN
 itmNouveau.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = COLOR_BRUN
   itmNouveau.ListSubItems(I_COL_SOUM_ID).ForeColor = COLOR_BRUN
 itmNouveau.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = COLOR_BRUN
   itmNouveau.ListSubItems(I_COL_SOUM_MONTAGE).ForeColor = COLOR_BRUN
 itmNouveau.ListSubItems(I_COL_SOUM_NO_SEQUENTIEL).ForeColor = COLOR_BRUN
9  itmNouveau.ListSubItems(I_COL_SOUM_NOM_COMMANDE).ForeColor = COLOR_BRUN
 itmNouveau.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_BRUN
 itmNouveau.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = COLOR_BRUN
 itmNouveau.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = COLOR_BRUN
 itmNouveau.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = COLOR_BRUN
 itmNouveau.ListSubItems(I_COL_SOUM_TEMPS).ForeColor = COLOR_BRUN
 itmNouveau.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = COLOR_BRUN
 End If

1 Call CalculerTempsFabrication

 'Calcul des prix
 Call CalculerPrix
 
 'On cache le ListView
1 frafournisseur.Visible = False

 m_bPieceInutile = False
 
 'Resélectionne le premier élément du listview
1 If lvwSoumission.ListItems.count > 0 Then
10  Call Deselect

10  lvwSoumission.ListItems(1).Selected = True
10  End If
10  Else
10  Call MsgBox("Quantité trop grande!", vbOKOnly, "Erreur")
10  End If
109End If

10  Exit Sub

Oups:

110 wOups "frmProjSoumElec", "ChoisirFournisseurMateriel", Err, Err.number, Err.Description
End Sub

Private Sub lvwFournisseur_DblClick()

 On Error GoTo Oups

 If m_bPieceInutile = True Or m_bChangementFRS = True Then
 Call ChoisirFournisseurMateriel
 Else
 Call ChoisirFournisseur
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "lvwFournisseur_DblClick", Err, Err.number, Err.Description
End Sub

Private Sub lvwPieces_DblClick()

 On Error GoTo Oups

 m_bPieceInutile = False
 m_bRecherchePiece = False
 m_bChangementFRS = False
 'On affiche lvwFournisseur selon la pièce choisie
 'Rempli les fournisseurs de la pièce choisie
 Call AfficherListeFournisseurs
 
 'si le listview n'est pas vide
 If lvwfournisseur.ListItems.count = 1 Then
 If MsgBox("Il n'y a aucun fournisseur pour cette pièce!" & vbNewLine & "Voulez-vous en ajouter?", vbYesNo, "Erreur") = vbYes Then
 Screen.MousePointer = vbHourglass
 
 'On ouvre le catalogue sur cet enregistrement
 Call FrmCatalogueElec.AfficherForm(cmbPieces.Text, lvwPieces.SelectedItem.SubItems(I_COL_PIECES_MANUFACT), lvwPieces.SelectedItem.SubItems(I_COL_PIECES_NO_ITEM))
 
 Screen.MousePointer = vbDefault
 End If
  End If

  Exit Sub

Oups:

  wOups "frmProjSoumElec", "lvwPieces_DblClick", Err, Err.number, Err.Description
End Sub

Private Sub AfficherListeFournisseurs()

 On Error GoTo Oups

 'Méthode qui sert à afficher la liste des fournisseurs
 'Affiche le frame seulement s'il y a des items dans le ListView
 Call RemplirListViewFournisseur
 
 If lvwfournisseur.ListItems.count > 1 Then
 If m_bRecherchePiece = True Then
 fraPieceTrouve.Visible = False
 End If

 frafournisseur.Visible = True
 Call lvwfournisseur.SetFocus
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "AfficherListeFournisseurs", Err, Err.number, Err.Description
End Sub

Private Sub lvwSoumission_KeyDown(KeyCode As Integer, Shift As Integer)

 On Error GoTo Oups

 'Si le ListView n'est pas vide
 If lvwSoumission.ListItems.count > 0 Then
 If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor = -2147483640 Then
 lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PIECE).ForeColor = vbBlack
 End If

 If Shift = vbCtrlMask Then
 If KeyCode = vbKeyF Then
 m_sTexteRecherche = InputBox("Quel est la pièce à rechercher?")

 If Trim$(m_sTexteRecherche) <> vbNullString Then
 Call RechercherPieceListViewSoumission(m_sTexteRecherche)
 End If
  Else
  If KeyCode = vbKeyN Then
  If Trim$(m_sTexteRecherche) <> "" Then
  Call RechercherPieceListViewSoumission(m_sTexteRecherche)
  End If
  Else
 'S'il n'est pas en mode affichage
  If m_bModeAffichage = False Then
 'Si ce n'est pas une sous-section
  If KeyCode = vbKeyC Then
 Call CopierPiece
 Else
 If KeyCode = vbKeyV Then
 Call CollerPiece
 End If
 End If
 End If
 End If
 End If
 Else
 'S'il n'est pas en mode affichage
 If m_bModeAffichage = False Then
 'Si ce n'est pas une section
 If lvwSoumission.SelectedItem.Tag <> vbNullString Then
 'Si ce n'est pas une sous-section
 If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) <> vbNullString Then
 'Si la touche pesée est Delete
 If KeyCode = vbKeyDelete Then
 Call EffacerItemListViewSoumission
 Else
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyN Then
 If m_eType = TYPE_PROJET Then
 If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) <> "Texte" And lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) <> "Text" Then
1  If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PROFIT).Tag <> "EXTRA" Then
 If KeyCode = vbKeyReturn Then
 Call FacturerDate
 Else
 Call FacturerNC
 End If
 Else
 Call MsgBox("Cette commande doit être faite dans le projet " & lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PROVENANCE), vbOKOnly, "Erreur")
 End If
 End If
 End If
 Else
 If KeyCode = vbKeyI Then
 If m_eType = TYPE_PROJET Then
 If lvwSoumission.SelectedItem.ListSubItems(I_COL_SOUM_PROFIT).Tag <> "EXTRA" Then
 lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_ID) = InputBox("Quel est l'ID", , lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_ID))
 Else
 Call MsgBox("Cette commande doit être faite dans le projet " & lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PROVENANCE), vbOKOnly, "Erreur")
 End If
 End If
 End If
 End If
 End If
 End If
 End If
 End If
 End If
End If

Exit Sub

Oups:

wOups "frmProjSoumElec", "lvwSoumission_KeyDown", Err, Err.number, Err.Description
End Sub

Private Sub FacturerDate()

 On Error GoTo Oups

 Dim iCompteur As Integer

 For iCompteur = 1 To lvwSoumission.ListItems.count
 'Si c'est pas une section
 If lvwSoumission.ListItems(iCompteur).Tag <> "" Then
 'Si c'est pas une sous-section
 If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> "" Then
 'Si la pièce est sélectionnée
 If lvwSoumission.ListItems(iCompteur).Selected = True Then
 'Si il y a une date d'écrit
 If Left$(lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_FACTURATION), 2) = "F-" Then
 'On l'enlève
 lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_FACTURATION) = ""
 Else
 'Si il n'y a rien
 If Trim$(lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_FACTURATION)) = "" Then
 'On ajoute la date de facturation
 lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_FACTURATION) = "F-" & txtDateFacturation.Text
  End If
  End If
  End If
  End If
  End If
  Next

  Exit Sub

Oups:

  wOups "frmProjSoumElec", "FacturerDate", Err, Err.number, Err.Description
End Sub

Private Sub FacturerNC()

 On Error GoTo Oups

 Dim iCompteur As Integer

 For iCompteur = 1 To lvwSoumission.ListItems.count
 'Si c'est pas une section
 If lvwSoumission.ListItems(iCompteur).Tag <> "" Then
 'Si c'est pas une sous-section
 If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> "" Then
 'Si la pièce est sélectionnée
 If lvwSoumission.ListItems(iCompteur).Selected = True Then
 'Si il y a NC d'écrit
 If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_FACTURATION) = "NC" Then
 'On l'enlève
 lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_FACTURATION) = ""
 Else
 'Si il n'y a rien
 If Trim$(lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_FACTURATION)) = "" Then
 'On ajoute la date de facturation
 lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_FACTURATION) = "NC"
  End If
  End If
  End If
  End If
  End If
  Next

  Exit Sub

Oups:

  wOups "frmProjSoumElec", "FacturerNC", Err, Err.number, Err.Description
End Sub

Private Sub RechercherPieceListViewSoumission(ByVal sTexte As String)
 
 On Error GoTo Oups

 Dim iCompteur As Integer
 Dim iSelected As Integer
 Dim bTrouve As Boolean

 If lvwSoumission.SelectedItem.Index = 1 Then
 iSelected = 1
 Else
 If lvwSoumission.SelectedItem.Index + 1 > lvwSoumission.ListItems.count Then
 iSelected = 1
 Else
 iSelected = lvwSoumission.SelectedItem.Index + 1
  End If
  End If

  For iCompteur = iSelected To lvwSoumission.ListItems.count
  If InStr(1, UCase(lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE)), UCase(sTexte)) > 0 Then
  Call lvwSoumission.SetFocus

  Call Deselect

  lvwSoumission.ListItems(iCompteur).Selected = True

  Call lvwSoumission.SelectedItem.EnsureVisible

 bTrouve = True

Exit For
 End If
Next

If bTrouve = False Then
 For iCompteur = 1 To iSelected - 1
 If InStr(1, UCase(lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE)), UCase(sTexte)) > 0 Then
 Call lvwSoumission.SetFocus

 Call Deselect

 lvwSoumission.ListItems(iCompteur).Selected = True

 Call lvwSoumission.SelectedItem.EnsureVisible

 bTrouve = True

 Exit For
 End If
 Next
End If

 If bTrouve = False Then
 Call MsgBox("Aucun enregistrement trouvé!", vbOKOnly, "Erreur")
 End If

1  Exit Sub

Oups:

 wOups "frmProjSoumElec", "RechercherPieceListViewSoumission", Err, Err.number, Err.Description
End Sub

Private Sub EffacerItemListViewSoumission()

 On Error GoTo Oups

 Dim bSeulSS As Boolean 'Pour savoir si c'est le seul enr. dans la sous-section
 Dim bSeulS As Boolean 'Pour savoir si c'est le seul enr. dans la section
 Dim iIndex As Integer
 Dim itmPrecedent As ListItem
 Dim itmSuivant As ListItem
 Dim iCompteur As Integer
 Dim sMessage As String
 Dim iNbreSelected As Integer
 Dim bSupprimer As Boolean
 Dim bPermission As Boolean

  For iCompteur = 1 To lvwSoumission.ListItems.count
  If lvwSoumission.ListItems(iCompteur).Selected = True Then
  iNbreSelected = iNbreSelected + 1

  If iNbreSelected > 1 Then
  Exit For
  End If
  End If
  Next

10 If iNbreSelected > 1 Then
1 sMessage = "Voulez-vous vraiment effacer ces pièces?"
Else
 sMessage = "Voulez-vous vraiment effacer cette pièce?"
End If

If m_eType = TYPE_SOUMISSION Then
 bPermission = True
Else
 If iNbreSelected > 1 Then
 bPermission = True
 Else
 For iCompteur = 1 To lvwSoumission.ListItems.count
 If lvwSoumission.ListItems(iCompteur).Selected = True Then
 Exit For
 End If
 Next

 If lvwSoumission.ListItems(iCompteur).ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_NOIR Or _
 lvwSoumission.ListItems(iCompteur).ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_BRUN Or _
 lvwSoumission.ListItems(iCompteur).ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_MAGENTA Then
 bPermission = True
 End If
1  End If
 End If

 iCompteur = 1

If bPermission = True Then
 If MsgBox(sMessage, vbYesNo) = vbYes Then
 Do While iCompteur <= lvwSoumission.ListItems.count
 bSupprimer = False
 bSeulS = False
 bSeulSS = False

 If lvwSoumission.ListItems(iCompteur).Selected = True Then
 'Si ce n'est pas une section
 If lvwSoumission.ListItems(iCompteur).Tag <> vbNullString Then
 'Si c'est une sous-section
 If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) <> vbNullString Then
 If m_eType = TYPE_SOUMISSION Then
 bSupprimer = True
 Else
 If lvwSoumission.ListItems(iCompteur).ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_NOIR Or _
 lvwSoumission.ListItems(iCompteur).ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_BRUN Or _
 lvwSoumission.ListItems(iCompteur).ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_MAGENTA Then
 bSupprimer = True
 End If
 End If

 If bSupprimer = True Then
 If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PROFIT) = "" Then
 lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PROFIT) = " "
 End If

 If lvwSoumission.ListItems(iCompteur).ListSubItems(I_COL_SOUM_PROFIT).Tag <> "EXTRA" Then
 If m_bModeAjout = False Then
 Call AjouterSuppressionCollection(iCompteur)
 End If
  
 iIndex = iCompteur
  
 'Il faut vérifier si c'est le seul enregistrement de la section. Si c'est le cas
 'Il faut effacer la section en meme temps
 'Si l'item sélectionné est le dernier enregistrement
 If iIndex = lvwSoumission.ListItems.count Then
 'Si l'élément d'en haut est une sous-section
 If lvwSoumission.ListItems(iIndex - 1).ListSubItems(I_COL_SOUM_PIECE) = vbNullString Then
 'Il est le seul dans la sous-section
 bSeulSS = True
 
 'Il faut maintenant vérifier si il est le seul dans la section
 If lvwSoumission.ListItems(iIndex - 2).Tag = vbNullString Then
 'Il est le seul enr. dans la section
 bSeulS = True
 End If
 End If
 Else
 Set itmPrecedent = lvwSoumission.ListItems(iIndex - 1)
 Set itmSuivant = lvwSoumission.ListItems(iIndex + 1)
 
 'Si l'élément précedent est une sous-section et le suivant est une sous-section ou une section
 If itmPrecedent.ListSubItems(I_COL_SOUM_PIECE).Tag = vbNullString And (itmSuivant.Tag = vbNullString Or (itmSuivant.Tag <> vbNullString And itmSuivant.ListSubItems(I_COL_SOUM_PIECE).Tag = vbNullString)) Then
 'C'est le seul dans la sous-section
 bSeulSS = True
 
 'Si les éléments avant et après sont des sections
 If lvwSoumission.ListItems(iIndex - 2).Tag = vbNullString And itmSuivant.Tag = vbNullString Then
4 bSeulS = True
4 End If
4 End If
4 End If
 
4 Call lvwSoumission.ListItems.Remove(iIndex)
 
 'Si c'est le seul dans la sous-section, on efface la sous-section
4 If bSeulSS = True Then
4 Call lvwSoumission.ListItems.Remove(iIndex - 1)

4 iCompteur = iCompteur - 1
4 End If
 
 'Si c'est le seul dans la section, on efface la section
4 If bSeulS = True Then
4 Call lvwSoumission.ListItems.Remove(iIndex - 2)

4  iCompteur = iCompteur - 1
4  End If
 
 'On recalcule le temps mécanique
4  Call CalculerTempsFabrication

 'On recalcule les prix
4  Call CalculerPrix
4  Else
4  Call MsgBox("La pièce " & lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) & " doit être effacée dans le projet " & lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PROVENANCE), vbOKOnly, "Erreur")

4  iCompteur = iCompteur + 1
4  End If
50 Else
 Call MsgBox("La pièce " & lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) & " ne peut pas être supprimée!", vbOKOnly, "Erreur")

 iCompteur = iCompteur + 1
 End If
 Else
 iCompteur = iCompteur + 1
 End If
 Else
 iCompteur = iCompteur + 1
 End If
 Else
 iCompteur = iCompteur + 1
5  End If
5  Loop
5  Else
 'Cette ligne sert seulement à ne pas déselectionner et repositionner à la ligne 1 si l'utilisateur
 'décide de ne pas supprimer.
 'Le nom de la variable n'est pas significatif dans ce cas, mais c'est celle-ci qui est utilisé pour
 'désélectionner et remettre à la ligne 1
5  bPermission = False
5  End If
5  End If
 
 'Il faut resélectionner le premier à la fin
5  If lvwSoumission.ListItems.count > 0 Then
5  If bPermission = True Then
60 Call Deselect

  lvwSoumission.ListItems(1).Selected = True
  End If
  End If

  Exit Sub

Oups:

  wOups "frmProjSoumElec", "EffacerItemListViewSoumission", Err, Err.number, Err.Description
End Sub

Private Sub AjouterSuppressionCollection(ByVal iIndex As Integer)

 On Error GoTo Oups

 If lvwSoumission.ListItems(iIndex).SubItems(I_COL_SOUM_PIECE) <> "Texte" And lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) <> "Text" Then
 Call m_collQteSupp.Add(Replace(lvwSoumission.ListItems(iIndex).Text, "*", ""))
 Call m_collNoItemSupp.Add(lvwSoumission.ListItems(iIndex).SubItems(I_COL_SOUM_PIECE))
 Call m_collDateSupp.Add(ConvertDate(Date))
 Call m_collHeureSupp.Add(Time)
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "AjouterSuppressionCollection", Err, Err.number, Err.Description
End Sub

Private Sub EnregistrerSuppression()

 On Error GoTo Oups

 Dim rstBavard As ADODB.Recordset
 Dim rstEmploye As ADODB.Recordset
 Dim iNoEmploye As Integer
 Dim iCompteur As Integer

 Set rstBavard = New ADODB.Recordset
 Set rstEmploye = New ADODB.Recordset

 Call rstBavard.Open("SELECT * FROM GrbBavardSuppression", g_connData, adOpenDynamic, adLockOptimistic)

 Call rstEmploye.Open("SELECT noEmploye FROM GrbEmployés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)

 iNoEmploye = rstEmploye.Fields("noEmploye")

 Call rstEmploye.Close
  Set rstEmploye = Nothing

  For iCompteur = 1 To m_collNoItemSupp.count
  Call rstBavard.AddNew

  rstBavard.Fields("IDUser") = iNoEmploye
  rstBavard.Fields("NoProjsoum") = txtNoProjSoum.Text
  rstBavard.Fields("Type") = "E"
  rstBavard.Fields("Qté") = m_collQteSupp(iCompteur)
  rstBavard.Fields("No Item") = m_collNoItemSupp(iCompteur)
rstBavard.Fields("Date") = m_collDateSupp(iCompteur)
1 rstBavard.Fields("Heure") = m_collHeureSupp(iCompteur)

 Call rstBavard.Update
Next

Call rstBavard.Close
Set rstBavard = Nothing

Exit Sub

Oups:

wOups "frmProjSoumElec", "EnregistrerSuppression", Err, Err.number, Err.Description
End Sub

Private Sub mvwDateRequise_GotFocus()
 
 On Error GoTo Oups

 m_bMonthViewHasFocus = True

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "mvwDateRequise_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub tmrTemps_Timer()
 
 On Error GoTo Oups

 If lblPasTemps.Visible = True Then
 lblPasTemps.Visible = False
 Else
 lblPasTemps.Visible = True
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "tmrTemps_Timer", Err, Err.number, Err.Description
End Sub

Private Sub txtCheminPhotos_KeyDown(KeyCode As Integer, Shift As Integer)

 On Error GoTo Oups

 If m_eMode = MODE_AJOUT_MODIF Then
 If KeyCode <> vbKeyBack And KeyCode <> vbKeyDelete Then
 KeyCode = 0
 End If
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "txtCheminPhotos_KeyDown", Err, Err.number, Err.Description
End Sub

Private Sub txtPrixManuel_Change()

 On Error GoTo Oups
 
 'Si le texte change, il faut recalculer les prix
 Call CalculerPrix

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "txtManuel_Change", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnulerPrix_Click()

 On Error GoTo Oups

 fraPrixPiece.Visible = False

 m_bMauvaisPrix = False

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "cmdAnnulerPrix_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdOKPrix_Click()
 'Écrit les prix dans le ListView
 On Error GoTo Oups

 Dim rstConfig As ADODB.Recordset
 Dim itmSoum As ListItem
 Dim itmAvant As ListItem
 Dim bPrixSpecial As Boolean
 Dim iCompteur As Integer
 Dim lColor As Long
 Dim sQuantite As String
 Dim sPiece As String
 Dim sTauxUSA As String
 Dim sTauxSPA As String

  Set rstConfig = New ADODB.Recordset

  Call rstConfig.Open("SELECT TauxAmericain, TauxEspagnol FROM GrbConfig", g_connData, adOpenForwardOnly, adLockReadOnly)

  sTauxUSA = rstConfig.Fields("TauxAmericain")
  sTauxSPA = rstConfig.Fields("TauxEspagnol")

  Call rstConfig.Close
  Set rstConfig = Nothing

  If m_bMauvaisPrix = False Then
  If cmbfrs.ListIndex = -1 Then
 Call MsgBox("Vous devez choisir un fournisseur!", vbOKOnly, "Erreur")
 
Exit Sub
 End If
End If
 
If Trim$(txtPrixList.Text) = vbNullString Then
 If Trim$(txtPrixSpecial.Text) = vbNullString Then
 Call MsgBox("Vous devez remplir le prix listé!", vbOKOnly, "Erreur")

 Exit Sub
 End If
End If
 
If Trim$(txtPrixNet.Text) = vbNullString And Trim$(txtPrixSpecial.Text) = vbNullString Then
 Call MsgBox("Vous devez choisir un prix!", vbOKOnly, "Erreur")
 
Exit Sub
Else
 If Trim$(txtPrixNet.Text) <> vbNullString Then
 bPrixSpecial = False
 Else
 bPrixSpecial = True
 End If
1  End If

 If m_bMauvaisPrix = True Then
 sQuantite = InputBox("Quelle est la quantité!")

 If sQuantite <> "" Then
 If Not IsNumeric(sQuantite) Then
 Exit Sub
 End If
 Else
 Exit Sub
 End If

 Set itmAvant = lvwSoumission.ListItems(CInt(fraPrixPiece.Tag))
 Set itmSoum = lvwSoumission.ListItems.Add(CInt(fraPrixPiece.Tag) + 1)

 lColor = itmAvant.ListSubItems(I_COL_SOUM_PIECE).ForeColor

itmSoum.Checked = itmAvant.Checked

 'Quantité
 itmSoum.Text = "-" & itmAvant.Text

 'On met l'id de la section dans le tag du listItem
itmSoum.Tag = itmAvant.Tag

 'No d'item
 itmSoum.SubItems(I_COL_SOUM_PIECE) = itmAvant.SubItems(I_COL_SOUM_PIECE)

 'On met le nom de la sous-section dans le tag du no d'item
itmSoum.ListSubItems(I_COL_SOUM_PIECE).Tag = itmAvant.ListSubItems(I_COL_SOUM_PIECE).Tag

 'On met la description en francais dans la colonne et la description en anglais
 'dans le tag
 itmSoum.SubItems(I_COL_SOUM_DESCR) = itmAvant.SubItems(I_COL_SOUM_DESCR)
itmSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = itmAvant.ListSubItems(I_COL_SOUM_DESCR).Tag

 'On met le nom du fabricant dans la col-nne et l'ordre de la section dans le tag
 itmSoum.SubItems(I_COL_SOUM_MANUFACT) = itmAvant.SubItems(I_COL_SOUM_MANUFACT)
itmSoum.ListSubItems(I_COL_SOUM_MANUFACT).Tag = itmAvant.ListSubItems(I_COL_SOUM_MANUFACT).Tag

 'Prix listé
3 itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = itmAvant.SubItems(I_COL_SOUM_PRIX_LIST)

 itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = itmAvant.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag

 itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = itmAvant.SubItems(I_COL_SOUM_ESCOMPTE)

 itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = itmAvant.SubItems(I_COL_SOUM_PRIX_NET)

 itmSoum.ListSubItems(I_COL_SOUM_PRIX_NET).Tag = itmAvant.ListSubItems(I_COL_SOUM_PRIX_NET).Tag

 'On met le fournisseur dans la colonne et l'id dans le tag
 itmSoum.SubItems(I_COL_SOUM_DISTRIB) = itmAvant.SubItems(I_COL_SOUM_DISTRIB)
 itmSoum.ListSubItems(I_COL_SOUM_DISTRIB).Tag = itmAvant.ListSubItems(I_COL_SOUM_DISTRIB).Tag

 'Temps
 itmSoum.SubItems(I_COL_SOUM_TEMPS) = itmAvant.SubItems(I_COL_SOUM_TEMPS)

 'Si le temps n'est pas vide
 If Trim$(itmSoum.SubItems(I_COL_SOUM_TEMPS)) <> vbNullString Then
 'On calcul le temps * quantité pour la colonne montage
 itmSoum.SubItems(I_COL_SOUM_MONTAGE) = CDbl(Replace(itmSoum.SubItems(I_COL_SOUM_TEMPS), ".", ",")) * CDbl(Replace(itmSoum.Text, "*", vbNullString))
 Else
 itmSoum.SubItems(I_COL_SOUM_MONTAGE) = vbNullString
 End If

 'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
itmSoum.SubItems(I_COL_SOUM_TOTAL) = "-" & itmAvant.SubItems(I_COL_SOUM_TOTAL)

 'Pour le profit, c'est le prix total - (prix net * quantité)
 itmSoum.SubItems(I_COL_SOUM_PROFIT) = "-" & itmAvant.SubItems(I_COL_SOUM_PROFIT)

 'Ajout de l'enregistrement avec le nouveau prix
Set itmSoum = lvwSoumission.ListItems.Add(CInt(fraPrixPiece.Tag) + 2)

 itmSoum.Checked = itmAvant.Checked

 'Quantité
 itmSoum.Text = sQuantite

 'On met l'id de la section dans le tag du listItem
 itmSoum.Tag = itmAvant.Tag

 'No d'item
itmSoum.SubItems(I_COL_SOUM_PIECE) = itmAvant.SubItems(I_COL_SOUM_PIECE)

4 itmSoum.ListSubItems(I_COL_SOUM_PIECE).ForeColor = lColor

 'On met le nom de la sous-section dans le tag du no d'item
4 itmSoum.ListSubItems(I_COL_SOUM_PIECE).Tag = itmAvant.ListSubItems(I_COL_SOUM_PIECE).Tag

 'On met la description en francais dans la colonne et la description en anglais
 'dans le tag
4 itmSoum.SubItems(I_COL_SOUM_DESCR) = itmAvant.SubItems(I_COL_SOUM_DESCR)
4 itmSoum.ListSubItems(I_COL_SOUM_DESCR).Tag = itmAvant.ListSubItems(I_COL_SOUM_DESCR).Tag

4 itmSoum.ListSubItems(I_COL_SOUM_DESCR).ForeColor = lColor

 'On met le nom du fabricant dans la col-nne et l'ordre de la section dans le tag
4 itmSoum.SubItems(I_COL_SOUM_MANUFACT) = itmAvant.SubItems(I_COL_SOUM_MANUFACT)
4 itmSoum.ListSubItems(I_COL_SOUM_MANUFACT).Tag = itmAvant.ListSubItems(I_COL_SOUM_MANUFACT).Tag

4 itmSoum.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = lColor

4 If bPrixSpecial = False Then
4 If optUSA.Value = True Then
4 itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixList.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
4  Else
4  If optSpain.Value = True Then
4  itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixList.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
4  Else
4  itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(txtPrixList.Text, MODE_ARGENT, 4)
4  End If
4  End If

4  itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = txtPrixList.Text
 
50 itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = lColor
 
 'Escompte
If mskEscompte.Text <> vbNullString Then
 itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion(mskEscompte.Text, MODE_POURCENT)
 Else
 itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion("0", MODE_POURCENT)
 End If

 itmSoum.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = lColor

 'Prix net
 If optUSA.Value = True Then
 itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixNet.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
 Else
 If optSpain.Value = True Then
 itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixNet.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
5  Else
5  itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(txtPrixNet.Text, MODE_ARGENT, 4)
5  End If
5  End If

5  itmSoum.ListSubItems(I_COL_SOUM_PRIX_NET).Tag = itmAvant.ListSubItems(I_COL_SOUM_PRIX_NET).Tag

5  itmSoum.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = lColor
5  Else
5  If optUSA.Value = True Then
60 itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
  Else
  If optSpain.Value = True Then
  itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
  Else
  itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(txtPrixSpecial.Text, MODE_ARGENT, 4)
  End If
  End If

  itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = txtPrixSpecial.Text

  itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = lColor

  itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion("0", MODE_POURCENT)

  itmSoum.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = lColor

6  If optUSA.Value = True Then
6  itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
6  Else
6  If optSpain.Value = True Then
6  itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
6  Else
6  itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(txtPrixSpecial.Text, MODE_ARGENT, 4)
6  End If
70 End If

  itmSoum.ListSubItems(I_COL_SOUM_PRIX_NET).Tag = itmAvant.ListSubItems(I_COL_SOUM_PRIX_NET).Tag

  itmSoum.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = lColor
  End If

 'On met le fournisseur dans la colonne et l'id dans le tag
  itmSoum.SubItems(I_COL_SOUM_DISTRIB) = itmAvant.SubItems(I_COL_SOUM_DISTRIB)
  itmSoum.ListSubItems(I_COL_SOUM_DISTRIB).Tag = itmAvant.ListSubItems(I_COL_SOUM_DISTRIB).Tag

  itmSoum.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = lColor
 
 'Temps
  itmSoum.SubItems(I_COL_SOUM_TEMPS) = itmAvant.SubItems(I_COL_SOUM_TEMPS)

  itmSoum.ListSubItems(I_COL_SOUM_TEMPS).ForeColor = lColor
 
 'Si le temps n'est pas vide
  If Trim$(itmSoum.SubItems(I_COL_SOUM_TEMPS)) <> vbNullString Then
 'On calcul le temps * quantité pour la colonne montage
  itmSoum.SubItems(I_COL_SOUM_MONTAGE) = CDbl(Replace(itmSoum.SubItems(I_COL_SOUM_TEMPS), ".", ",")) * CDbl(Replace(itmSoum.Text, "*", vbNullString))
  Else
   itmSoum.SubItems(I_COL_SOUM_MONTAGE) = vbNullString
   End If
 
7  itmSoum.ListSubItems(I_COL_SOUM_MONTAGE).ForeColor = lColor

 'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
7  itmSoum.SubItems(I_COL_SOUM_TOTAL) = Conversion(CStr(Round(Replace(itmSoum.Text, "*", vbNullString) * itmSoum.SubItems(I_COL_SOUM_PRIX_NET) * CSng(m_sProfit), 2)), MODE_ARGENT)
 
7  itmSoum.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = lColor

7  If optUSA.Value = True Then
7  itmSoum.ListSubItems(I_COL_SOUM_TOTAL).Tag = "USA"
7  Else
80 If optSpain.Value = True Then
  itmSoum.ListSubItems(I_COL_SOUM_TOTAL).Tag = "SPA"
  Else
  itmSoum.ListSubItems(I_COL_SOUM_TOTAL).Tag = "CAN"
  End If
  End If
 
 'Pour le profit, c'est le prix total - (prix net * quantité)
  itmSoum.SubItems(I_COL_SOUM_PROFIT) = Conversion(CStr(Round(itmSoum.SubItems(I_COL_SOUM_TOTAL) - (itmSoum.SubItems(I_COL_SOUM_PRIX_NET) * Replace(itmSoum.Text, "*", vbNullString)), 2)), MODE_ARGENT)

  itmSoum.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = lColor

  itmSoum.SubItems(I_COL_SOUM_COMMENTAIRE) = itmAvant.SubItems(I_COL_SOUM_COMMENTAIRE)
  itmSoum.ListSubItems(I_COL_SOUM_COMMENTAIRE).ForeColor = lColor

  If m_eType = TYPE_PROJET Then
  itmSoum.SubItems(I_COL_SOUM_DATE_COMMANDE) = itmAvant.SubItems(I_COL_SOUM_DATE_COMMANDE)
   itmSoum.ListSubItems(I_COL_SOUM_DATE_COMMANDE).ForeColor = lColor

   itmSoum.SubItems(I_COL_SOUM_DATE_REQUISE) = itmAvant.SubItems(I_COL_SOUM_DATE_REQUISE)
   itmSoum.ListSubItems(I_COL_SOUM_DATE_REQUISE).ForeColor = lColor

   itmSoum.SubItems(I_COL_SOUM_NOM_COMMANDE) = itmAvant.SubItems(I_COL_SOUM_NOM_COMMANDE)
8  itmSoum.ListSubItems(I_COL_SOUM_NOM_COMMANDE).ForeColor = lColor

8  itmSoum.SubItems(I_COL_SOUM_NO_SEQUENTIEL) = itmAvant.SubItems(I_COL_SOUM_NO_SEQUENTIEL)
8  itmSoum.ListSubItems(I_COL_SOUM_NO_SEQUENTIEL).ForeColor = lColor

8  itmSoum.SubItems(I_COL_SOUM_ID) = itmAvant.SubItems(I_COL_SOUM_ID)
90 itmSoum.ListSubItems(I_COL_SOUM_ID).ForeColor = lColor
 
  itmSoum.SubItems(I_COL_SOUM_FACTURATION) = itmAvant.SubItems(I_COL_SOUM_FACTURATION)

  If itmSoum.SubItems(I_COL_SOUM_FACTURATION) <> "" Then
  itmSoum.ListSubItems(I_COL_SOUM_FACTURATION).Tag = itmAvant.ListSubItems(I_COL_SOUM_FACTURATION)
  End If

  itmSoum.ListSubItems(I_COL_SOUM_FACTURATION).ForeColor = lColor
  End If

  If itmAvant.SubItems(I_COL_SOUM_COMMENTAIRE) <> "" Then
  itmSoum.ListSubItems(I_COL_SOUM_COMMENTAIRE).ForeColor = itmAvant.ListSubItems(I_COL_SOUM_COMMENTAIRE).ForeColor

  itmAvant.ListSubItems(I_COL_SOUM_COMMENTAIRE).ForeColor = vbBlack
  End If

  If m_eType = TYPE_PROJET Then
 If itmAvant.SubItems(I_COL_SOUM_DATE_COMMANDE) <> "" Then
   itmAvant.ListSubItems(I_COL_SOUM_DATE_COMMANDE).ForeColor = vbBlack
 End If

   If itmAvant.SubItems(I_COL_SOUM_DATE_REQUISE) <> "" Then
 itmAvant.ListSubItems(I_COL_SOUM_DATE_REQUISE).ForeColor = vbBlack
   End If

 If itmAvant.SubItems(I_COL_SOUM_NOM_COMMANDE) <> "" Then
9  itmAvant.ListSubItems(I_COL_SOUM_NOM_COMMANDE).ForeColor = vbBlack
 End If

10 If itmAvant.SubItems(I_COL_SOUM_NO_SEQUENTIEL) <> "" Then
 itmAvant.ListSubItems(I_COL_SOUM_NO_SEQUENTIEL).ForeColor = vbBlack
1 End If

 If itmAvant.SubItems(I_COL_SOUM_FACTURATION) <> "" Then
 itmAvant.ListSubItems(I_COL_SOUM_FACTURATION).ForeColor = vbBlack
 End If

1 If itmAvant.SubItems(I_COL_SOUM_ID) <> "" Then
 itmAvant.ListSubItems(I_COL_SOUM_ID).ForeColor = vbBlack
1 End If

 If itmAvant.SubItems(I_COL_SOUM_PROVENANCE) <> "" Then
 itmAvant.ListSubItems(I_COL_SOUM_PROVENANCE).ForeColor = vbBlack
10  End If
10  End If

10  itmSoum.ListSubItems(I_COL_SOUM_DESCR).ForeColor = itmAvant.ListSubItems(I_COL_SOUM_DESCR).ForeColor
10  itmSoum.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = itmAvant.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor
10  itmSoum.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = itmAvant.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor
10  itmSoum.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = itmAvant.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor
10  itmSoum.ListSubItems(I_COL_SOUM_MONTAGE).ForeColor = itmAvant.ListSubItems(I_COL_SOUM_MONTAGE).ForeColor
10  itmSoum.ListSubItems(I_COL_SOUM_PIECE).ForeColor = itmAvant.ListSubItems(I_COL_SOUM_PIECE).ForeColor
110 itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = itmAvant.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor
11 itmSoum.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = itmAvant.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor
1 itmSoum.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = itmAvant.ListSubItems(I_COL_SOUM_PROFIT).ForeColor
1 itmSoum.ListSubItems(I_COL_SOUM_TEMPS).ForeColor = itmAvant.ListSubItems(I_COL_SOUM_TEMPS).ForeColor
1 itmSoum.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = itmAvant.ListSubItems(I_COL_SOUM_TOTAL).ForeColor

1 itmAvant.ListSubItems(I_COL_SOUM_DESCR).ForeColor = vbBlack
1 itmAvant.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = vbBlack
1 itmAvant.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = vbBlack
1 itmAvant.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = vbBlack
1 itmAvant.ListSubItems(I_COL_SOUM_MONTAGE).ForeColor = vbBlack
1 itmAvant.ListSubItems(I_COL_SOUM_PIECE).ForeColor = vbBlack
1 itmAvant.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = vbBlack
11  itmAvant.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = vbBlack
1 itmAvant.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = vbBlack
 itmAvant.ListSubItems(I_COL_SOUM_TEMPS).ForeColor = vbBlack
1 itmAvant.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = vbBlack

 Call CalculerTempsFabrication

 'Resélectionne le premier élément du listview
1 If lvwSoumission.ListItems.count > 0 Then
 Call Deselect

11  lvwSoumission.ListItems(1).Selected = True
 End If
 
1 m_bMauvaisPrix = False

1 cmbfrs.Locked = False

1 Call lvwSoumission.Refresh
12 Else
1 sPiece = lvwSoumission.ListItems(CInt(fraPrixPiece.Tag)).SubItems(I_COL_SOUM_PIECE)

1 For iCompteur = 1 To lvwSoumission.ListItems.count
1 If lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) = sPiece And lvwSoumission.ListItems(iCompteur).ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_MAGENTA Then
1 Set itmSoum = lvwSoumission.ListItems(iCompteur)

1 itmSoum.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_NOIR
1 itmSoum.ListSubItems(I_COL_SOUM_DESCR).ForeColor = COLOR_NOIR
1 itmSoum.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = COLOR_NOIR
1 itmSoum.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = COLOR_NOIR
1 itmSoum.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = COLOR_NOIR
1 itmSoum.ListSubItems(I_COL_SOUM_MONTAGE).ForeColor = COLOR_NOIR
1 itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = COLOR_NOIR
1 itmSoum.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = COLOR_NOIR
1 itmSoum.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = COLOR_NOIR
1 itmSoum.ListSubItems(I_COL_SOUM_TEMPS).ForeColor = COLOR_NOIR
1 itmSoum.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = COLOR_NOIR

1 If itmSoum.SubItems(I_COL_SOUM_COMMENTAIRE) <> "" Then
1 itmSoum.ListSubItems(I_COL_SOUM_COMMENTAIRE).ForeColor = COLOR_NOIR
1 End If

1 Call lvwSoumission.Refresh
 
1 If bPrixSpecial = False Then
 'Prix listé
1 If optUSA.Value = True Then
1 itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixList.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
1 Else
1 If optSpain.Value = True Then
1 itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixList.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
1 Else
1 itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(txtPrixList.Text, MODE_ARGENT, 4)
1 End If
1 End If

1 itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = txtPrixList.Text
 
 'Escompte
1 If mskEscompte.Text <> vbNullString Then
1 itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion(mskEscompte.Text, MODE_POURCENT)
1 Else
1 itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion("0", MODE_POURCENT)
1 End If

 'Prix net
1 If optUSA.Value = True Then
14 itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixNet.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
14 Else
14 If optSpain.Value = True Then
14 itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixNet.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
14 Else
14 itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(txtPrixNet.Text, MODE_ARGENT, 4)
14 End If
14 End If
14 Else
14 If optUSA.Value = True Then
14 itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
14  Else
14  If optSpain.Value = True Then
14  itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
14  Else
14  itmSoum.SubItems(I_COL_SOUM_PRIX_LIST) = Conversion(txtPrixSpecial.Text, MODE_ARGENT, 4)
14  End If
14  End If

14  itmSoum.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = txtPrixSpecial.Text
 
150 itmSoum.SubItems(I_COL_SOUM_ESCOMPTE) = Conversion("0", MODE_POURCENT)

1 If optUSA.Value = True Then
 itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
 Else
 If optSpain.Value = True Then
 itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
 Else
 itmSoum.SubItems(I_COL_SOUM_PRIX_NET) = Conversion(txtPrixSpecial.Text, MODE_ARGENT, 4)
 End If
 End If
 End If

 'On met le fournisseur dans la colonne et l'id dans le tag
 itmSoum.SubItems(I_COL_SOUM_DISTRIB) = cmbfrs.LIST(cmbfrs.ListIndex)
 
15  itmSoum.ListSubItems(I_COL_SOUM_DISTRIB).Tag = cmbfrs.ItemData(cmbfrs.ListIndex)
 
 'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
15  itmSoum.SubItems(I_COL_SOUM_TOTAL) = Conversion(CStr(Round(Replace(itmSoum.Text, "*", "") * itmSoum.SubItems(I_COL_SOUM_PRIX_NET) * CSng(m_sProfit), 2)), MODE_ARGENT)

15  If optUSA.Value = True Then
15  itmSoum.ListSubItems(I_COL_SOUM_TOTAL).Tag = "USA"
15  Else
15  If optSpain.Value = True Then
15  itmSoum.ListSubItems(I_COL_SOUM_TOTAL).Tag = "SPA"
15  Else
160 itmSoum.ListSubItems(I_COL_SOUM_TOTAL).Tag = "CAN"
 End If
 End If
 
 'Pour le profit, c'est le prix total - (prix net * quantité)
 itmSoum.SubItems(I_COL_SOUM_PROFIT) = Conversion(CStr(Round(itmSoum.SubItems(I_COL_SOUM_TOTAL) - (itmSoum.SubItems(I_COL_SOUM_PRIX_NET) * Replace(itmSoum.Text, "*", "")), 2)), MODE_ARGENT)
 End If
 Next
1  End If

1  Call ModifierPrixCatalogue

1  fraPrixPiece.Visible = False

1  Call CalculerPrix

1  Exit Sub

Oups:

1  wOups "frmProjSoumElec", "cmdOKPrix_Click", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboFournisseur()

 On Error GoTo Oups

 Dim rstFRS As ADODB.Recordset
 Dim iCompteur As Integer
 Dim bExiste As Boolean

 Set rstFRS = New ADODB.Recordset

 'Il faut vider le combo avant de le remplir
 Call cmbfrs.Clear

 Call rstFRS.Open("SELECT GrbPiecesFRS.*, GrbFournisseur.NomFournisseur FROM GrbPiecesFRS INNER JOIN GrbFournisseur ON GrbPiecesFRS.IDFRS = GrbFournisseur.IDFRS WHERE PIECE = '" & Replace(lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE), "'", "''") & "' ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)

 'Tant que ce n'est pas la fin des enregistrements
 Do While Not rstFRS.EOF
 bExiste = False

 For iCompteur = 0 To cmbfrs.ListCount - 1
 If cmbfrs.ItemData(iCompteur) = rstFRS.Fields("IDFRS") Then
  bExiste = True

  Exit For
  End If
  Next

  If bExiste = False Then
  Call cmbfrs.AddItem(rstFRS.Fields("NomFournisseur"))

  cmbfrs.ItemData(cmbfrs.newIndex) = rstFRS.Fields("IDFRS")
  End If

Call rstFRS.MoveNext
Loop

Call rstFRS.Close
Set rstFRS = Nothing

Exit Sub

Oups:

wOups "frmProjSoumElec", "RemplirComboFournisseur", Err, Err.number, Err.Description
End Sub

Private Sub txtPrixList_LostFocus()

 On Error GoTo Oups

 If txtPrixList.Text <> vbNullString Then
 txtPrixList.Text = Replace(txtPrixList, ".", ",")
 
 If IsNumeric(txtPrixList.Text) Then
 Call CalculerPrixNet
 Else
 Call MsgBox("Valeur non numérique!", vbOKOnly, "Erreur")
 txtPrixList.Text = vbNullString
 End If
 End If

 Exit Sub

Oups:

  wOups "frmProjSoumElec", "txtPrixList_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtPrixNet_Change()

 On Error GoTo Oups

 'Quand le contenu du prix net change
 
 'Si la longueur du texte écrit est plus grand que 0
 If Len(txtPrixNet.Text) > 0 Then
 'On vide le prix spécial et on le désactive
 txtPrixSpecial.Text = vbNullString
 txtPrixSpecial.Enabled = False
 Else
 'Sinon, on active le prix spécial
 txtPrixSpecial.Enabled = True
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "txtPrixNet_Change", Err, Err.number, Err.Description

End Sub

Private Sub txtPrixNet_GotFocus()

 On Error GoTo Oups

 'Si le prix net prend le focus
 Call CalculerPrixNet

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "txtPrixNet_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub CalculerPrixNet()

 On Error GoTo Oups

 Dim dblEscompte As Double
 Dim dblPrix As Double
 
 'Si le prix net n'est pas barré.. ie.. si le prix spécial est vide
 If txtPrixNet.Locked = False Then
 mskEscompte.Text = Replace(mskEscompte.Text, "_", vbNullString)
 
 mskEscompte.Text = Replace(mskEscompte.Text, ".", ",")
 
 If mskEscompte.Text <> vbNullString Then
 dblEscompte = CDbl(mskEscompte.Text)
 Else
 dblEscompte = 0
 End If
 
  If txtPrixList.Text <> vbNullString Then
  dblPrix = CDbl(Replace(txtPrixList.Text, ".", ","))
  Else
  dblPrix = 0
  End If
 
 'Calcul du prix net
  txtPrixNet.Text = Round((dblPrix) * (1 - dblEscompte), 4)
 
  txtPrixNet.Text = Replace(txtPrixNet.Text, ".", ",")
  End If

10 Exit Sub

Oups:

wOups "frmProjSoumElec", "CalculerPrixNet", Err, Err.number, Err.Description
End Sub

Private Sub txtPrixNet_LostFocus()

 On Error GoTo Oups

 txtPrixNet.Text = Replace(txtPrixNet.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "txtPrixNet_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub ViderChamps_frs()

 On Error GoTo Oups

 'Vide les champs pieces
 txtPrixList.Text = vbNullString
 mskEscompte.Text = vbNullString
 txtPrixNet.Text = vbNullString
 txtPrixSpecial.Text = vbNullString
 
 optCAN.Value = True

 Call AfficherDrapeau

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "ViderChamps_frs", Err, Err.number, Err.Description
End Sub

Private Sub ModifierPrixCatalogue()
 'Enregistrement du prix de la pièce
 
 On Error GoTo Oups

 Dim rstPrix As ADODB.Recordset
 Dim dblPrixList As Double
 Dim dblEscompte As Double
 Dim dblPrixNet As Double
 
 If Trim$(txtPrixList.Text) <> "" Then
 dblPrixList = CDbl(txtPrixList.Text)
 Else
 dblPrixList = 0
 End If
 
 If mskEscompte.Text <> vbNullString Then
  dblEscompte = CDbl(mskEscompte.Text)
  Else
  dblEscompte = 0
  End If
 
  If Trim$(txtPrixNet.Text) <> "" Then
  dblPrixNet = CDbl(txtPrixNet.Text)
  Else
  dblPrixNet = CDbl(txtPrixSpecial.Text)
10 End If
  
Set rstPrix = New ADODB.Recordset
  
If txtPrixNet.Enabled = True Then
 'Ouverture du recordset
 Call rstPrix.Open("SELECT * FROM GrbPiecesFRS WHERE PIECE = '" & Replace(lvwSoumission.ListItems(CInt(fraPrixPiece.Tag)).SubItems(I_COL_SOUM_PIECE), "'", "''") & "' AND IDFRS = " & cmbfrs.ItemData(cmbfrs.ListIndex) & " AND PRIX_NET <> ''", g_connData, adOpenDynamic, adLockOptimistic)

 If rstPrix.EOF Then
 Call rstPrix.AddNew

 rstPrix.Fields("PIECE").Value = lvwSoumission.ListItems(CInt(fraPrixPiece.Tag)).SubItems(I_COL_SOUM_PIECE)
 rstPrix.Fields("IDFRS").Value = cmbfrs.ItemData(cmbfrs.ListIndex)
 End If

 rstPrix.Fields("PRIX_LIST") = dblPrixList
 rstPrix.Fields("ESCOMPTE") = dblEscompte
 rstPrix.Fields("PRIX_NET") = dblPrixNet
rstPrix.Fields("PRIX_SP") = ""
Else
 Call rstPrix.Open("SELECT * FROM GrbPiecesFRS WHERE PIECE = '" & Replace(lvwSoumission.ListItems(CInt(fraPrixPiece.Tag)).SubItems(I_COL_SOUM_PIECE), "'", "''") & "' AND IDFRS = " & cmbfrs.ItemData(cmbfrs.ListIndex) & " AND PRIX_SP <> ''", g_connData, adOpenDynamic, adLockOptimistic)

 If rstPrix.EOF Then
 Call rstPrix.AddNew

 rstPrix.Fields("PIECE").Value = lvwSoumission.ListItems(CInt(fraPrixPiece.Tag)).SubItems(I_COL_SOUM_PIECE)
 rstPrix.Fields("IDFRS").Value = cmbfrs.ItemData(cmbfrs.ListIndex)
1  End If

 rstPrix.Fields("PRIX_SP") = dblPrixNet
 rstPrix.Fields("PRIX_LIST") = ""
 rstPrix.Fields("ESCOMPTE") = ""
 rstPrix.Fields("PRIX_NET") = ""
End If

If optCAN.Value = True Then
 rstPrix.Fields("DeviseMonétaire") = "CAN"
Else
 If optUSA.Value = True Then
 rstPrix.Fields("DeviseMonétaire") = "USA"
 Else
 rstPrix.Fields("DeviseMonétaire") = "SPA"
End If
End If
 
2  rstPrix.Fields("Type") = "E"

rstPrix.Fields("ENTRER_PAR") = g_sInitiale

2  rstPrix.Fields("Date") = ConvertDate(Date)

Call rstPrix.Update
 
2  Call rstPrix.Close
Set rstPrix = Nothing
 
30 Exit Sub

Oups:

wOups "frmProjSoumElec", "ModifierPrixCatalogue", Err, Err.number, Err.Description
End Sub

Private Sub optCAN_Click()

 On Error GoTo Oups

 'dependant la devise, affiche le drapeau
 Call AfficherDrapeau

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "optCAN_Click", Err, Err.number, Err.Description
End Sub
 
Private Sub AfficherDrapeau()

 On Error GoTo Oups

 '''''''''''''''''''''''''''''
 'dependant la devise, affiche le drapeau
 '''''''''''''''''''''''''''''''''''''
 If optCAN.Value = True Then
 imgCanada.Visible = True
 imgEU.Visible = False
 imgSpain.Visible = False
 Else
 If optUSA.Value = True Then
 imgEU.Visible = True
 imgCanada.Visible = False
 imgSpain.Visible = False
 Else
  imgSpain.Visible = True
  imgCanada.Visible = False
  imgEU.Visible = False
  End If
  End If

  Exit Sub

Oups:

  wOups "frmProjSoumElec", "AfficherDrapeau", Err, Err.number, Err.Description
End Sub

Private Sub optSpain_Click()

 On Error GoTo Oups

 'dependant la devise, affiche le drapeau
 Call AfficherDrapeau

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "optSpain_Click", Err, Err.number, Err.Description
End Sub

Private Sub optUSA_Click()

 On Error GoTo Oups

 'Dépendant la devise, affiche le drapeau
 Call AfficherDrapeau

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "optUSA_Click", Err, Err.number, Err.Description
End Sub

Private Sub mskEscompte_GotFocus()

 On Error GoTo Oups

 'Quand le maskEdit prend le focus, on set le masque
 If mskEscompte.Enabled = True Then
 mskEscompte.mask = "0,####"
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "mskEscompte_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskEscompte_LostFocus()

 On Error GoTo Oups

 'Quand le maskEdit perd le focus, on enlève le mask
 mskEscompte.mask = vbNullString
 
 'Si le champs contient 0,____, c'est parce que rien n'a été entré
 If mskEscompte.Text = "0,____" Then
 'Donc, on le vide
 mskEscompte.Text = vbNullString
 End If
 
 Call CalculerPrixNet

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "mskEscompte_LostFocus", Err, Err.number, Err.Description
End Sub

Private Function VerifierSiOuvert(ByRef sUser As String) As Boolean
 'Vérifie si le projet ou la soumission n'est pas en modification
 'par un autre utilisateur sur un autre ordinateur
 On Error GoTo Oups

 Dim rstProjSoum As ADODB.Recordset
 Dim bModification As Boolean

 Set rstProjSoum = New ADODB.Recordset

 If m_eType = TYPE_PROJET Then
 Call rstProjSoum.Open("SELECT Modification, Par FROM GrbProjetElec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstProjSoum.Open("SELECT Modification, Par FROM GrbSoumissionElec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
 End If

 If rstProjSoum.Fields("Modification") = True Then
 sUser = rstProjSoum.Fields("Par")
  bModification = True
  Else
  sUser = ""
  bModification = False
  End If

  Call rstProjSoum.Close
  Set rstProjSoum = Nothing

  VerifierSiOuvert = bModification

10 Exit Function

Oups:

wOups "frmProjSoumElec", "VerifierSiOuvert", Err, Err.number, Err.Description
End Function

Private Sub OuvrirProjSoum(ByVal bOuvrir As Boolean)
 'Remplis ou vide les champs Modification et Par
 On Error GoTo Oups

 Dim rstProjSoum As ADODB.Recordset

 Set rstProjSoum = New ADODB.Recordset

 rstProjSoum.CursorLocation = adUseServer

 If m_eType = TYPE_PROJET Then
 Call rstProjSoum.Open("SELECT Modification, Par FROM GrbProjetElec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstProjSoum.Open("SELECT Modification, Par FROM GrbSoumissionElec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
 End If

 Do While Not rstProjSoum.EOF
 If bOuvrir = True Then
  rstProjSoum.Fields("Modification") = True
  rstProjSoum.Fields("Par") = g_sEmploye
  Else
  rstProjSoum.Fields("Modification") = False
  rstProjSoum.Fields("Par") = ""
  End If

  Call rstProjSoum.Update
 
  Call rstProjSoum.MoveNext
10 Loop

Call rstProjSoum.Close
Set rstProjSoum = Nothing

Exit Sub

Oups:

wOups "frmProjSoumElec", "OuvrirProjSoum", Err, Err.number, Err.Description
End Sub

Private Sub AnnulerCommande()

 On Error GoTo Oups

 Dim rstProjet As ADODB.Recordset
 Dim itmAvant As ListItem
 Dim itmAnnulation As ListItem
 Dim sExtra As String

 If Right$(txtNoProjSoum.Text, 2) >= "00" And Right$(txtNoProjSoum.Text, 2) <= "19" Then
 sExtra = InputBox("Dans quel extra l'annulation de commande doit être faite ? (2 chiffres seulement)")

 If Len(sExtra) <> 2 Then
 Call MsgBox("Format incorrect!", vbOKOnly, "Erreur")

 Exit Sub
 End If

  If Not IsNumeric(sExtra) Then
  Call MsgBox("L'extra doit être numérique!", vbOKOnly, "Erreur")

  Exit Sub
  End If

  If sExtra < 60 Or sExtra >   Then
  Call MsgBox("L'extra doit être entre 60 et 98!", vbOKOnly, "Erreur")

  Exit Sub
  End If

Set rstProjet = New ADODB.Recordset

1 Call rstProjet.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sExtra & "'", g_connData, adOpenDynamic, adLockOptimistic)

 If rstProjet.EOF Then
 Call MsgBox("Le projet " & Left$(txtNoProjSoum.Text, Len(txtNoProjSoum.Text) - 2) & sExtra & " n'existe pas!", vbOKOnly, "Erreur")

 Call rstProjet.Close
 Set rstProjet = Nothing

 Exit Sub
 Else
 Call rstProjet.Close
 Set rstProjet = Nothing
 End If
End If

1  Set itmAvant = lvwSoumission.SelectedItem
Set itmAnnulation = lvwSoumission.ListItems.Add(itmAvant.Index + 1)

 itmAnnulation.Checked = itmAvant.Checked

 'Quantité
itmAnnulation.Text = "-" & itmAvant.Text

 'On met l'id de la section dans le tag du listItem
 itmAnnulation.Tag = itmAvant.Tag

 'No d'item
itmAnnulation.SubItems(I_COL_SOUM_PIECE) = itmAvant.SubItems(I_COL_SOUM_PIECE)

 'On met le nom de la sous-section dans le tag du no d'item
 itmAnnulation.ListSubItems(I_COL_SOUM_PIECE).Tag = itmAvant.ListSubItems(I_COL_SOUM_PIECE).Tag

 'On met la description en francais dans la colonne et la description en anglais
 'dans le tag
1  itmAnnulation.SubItems(I_COL_SOUM_DESCR) = itmAvant.SubItems(I_COL_SOUM_DESCR)
 itmAnnulation.ListSubItems(I_COL_SOUM_DESCR).Tag = itmAvant.ListSubItems(I_COL_SOUM_DESCR).Tag

 'On met le nom du fabricant dans la col-nne et l'ordre de la section dans le tag
 itmAnnulation.SubItems(I_COL_SOUM_MANUFACT) = itmAvant.SubItems(I_COL_SOUM_MANUFACT)
itmAnnulation.ListSubItems(I_COL_SOUM_MANUFACT).Tag = itmAvant.ListSubItems(I_COL_SOUM_MANUFACT).Tag

 'Prix listé
itmAnnulation.SubItems(I_COL_SOUM_PRIX_LIST) = itmAvant.SubItems(I_COL_SOUM_PRIX_LIST)

itmAnnulation.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = itmAvant.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag

itmAnnulation.SubItems(I_COL_SOUM_ESCOMPTE) = itmAvant.SubItems(I_COL_SOUM_ESCOMPTE)

itmAnnulation.SubItems(I_COL_SOUM_PRIX_NET) = itmAvant.SubItems(I_COL_SOUM_PRIX_NET)

 'On met le fournisseur dans la colonne et l'id dans le tag
itmAnnulation.SubItems(I_COL_SOUM_DISTRIB) = itmAvant.SubItems(I_COL_SOUM_DISTRIB)
itmAnnulation.ListSubItems(I_COL_SOUM_DISTRIB).Tag = itmAvant.ListSubItems(I_COL_SOUM_DISTRIB).Tag

 'Temps
itmAnnulation.SubItems(I_COL_SOUM_TEMPS) = itmAvant.SubItems(I_COL_SOUM_TEMPS)

 'Si le temps n'est pas vide
If Trim$(itmAnnulation.SubItems(I_COL_SOUM_TEMPS)) <> vbNullString Then
 'On calcul le temps * quantité pour la colonne montage
 itmAnnulation.SubItems(I_COL_SOUM_MONTAGE) = CDbl(Replace(itmAnnulation.SubItems(I_COL_SOUM_TEMPS), ".", ",")) * CDbl(Replace(itmAnnulation.Text, "*", vbNullString))
2  Else
 itmAnnulation.SubItems(I_COL_SOUM_MONTAGE) = vbNullString
2  End If

 'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
itmAnnulation.SubItems(I_COL_SOUM_TOTAL) = "-" & itmAvant.SubItems(I_COL_SOUM_TOTAL)

 'Pour le profit, c'est le prix total - (prix net * quantité)
2  itmAnnulation.SubItems(I_COL_SOUM_PROFIT) = "-" & itmAvant.SubItems(I_COL_SOUM_PROFIT)

If Right$(txtNoProjSoum.Text, 2) >= "00" And Right$(txtNoProjSoum.Text, 2) <= "19" Then
 'Pour savoir lors de l'enregistremenet qu'il faut le lier avec un extra
itmAnnulation.ListSubItems(I_COL_SOUM_PROFIT).Tag = "ANNULATION " & sExtra
End If

30 itmAnnulation.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_VERT_FORET
itmAnnulation.ListSubItems(I_COL_SOUM_DESCR).ForeColor = COLOR_VERT_FORET
itmAnnulation.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = COLOR_VERT_FORET
itmAnnulation.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = COLOR_VERT_FORET
itmAnnulation.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = COLOR_VERT_FORET
itmAnnulation.ListSubItems(I_COL_SOUM_MONTAGE).ForeColor = COLOR_VERT_FORET
itmAnnulation.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = COLOR_VERT_FORET
itmAnnulation.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = COLOR_VERT_FORET
itmAnnulation.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = COLOR_VERT_FORET
itmAnnulation.ListSubItems(I_COL_SOUM_TEMPS).ForeColor = COLOR_VERT_FORET
itmAnnulation.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = COLOR_VERT_FORET
 
itmAnnulation.ListSubItems(I_COL_SOUM_PIECE).Bold = True
3  itmAnnulation.ListSubItems(I_COL_SOUM_DESCR).Bold = True
itmAnnulation.ListSubItems(I_COL_SOUM_DISTRIB).Bold = True
3  itmAnnulation.ListSubItems(I_COL_SOUM_ESCOMPTE).Bold = True
itmAnnulation.ListSubItems(I_COL_SOUM_MANUFACT).Bold = True
3  itmAnnulation.ListSubItems(I_COL_SOUM_MONTAGE).Bold = True
itmAnnulation.ListSubItems(I_COL_SOUM_PRIX_LIST).Bold = True
3  itmAnnulation.ListSubItems(I_COL_SOUM_PRIX_NET).Bold = True
 itmAnnulation.ListSubItems(I_COL_SOUM_PROFIT).Bold = True
40 itmAnnulation.ListSubItems(I_COL_SOUM_TEMPS).Bold = True
itmAnnulation.ListSubItems(I_COL_SOUM_TOTAL).Bold = True

4 itmAvant.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_VERT_FORET
4 itmAvant.ListSubItems(I_COL_SOUM_DESCR).ForeColor = COLOR_VERT_FORET
4 itmAvant.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = COLOR_VERT_FORET
4 itmAvant.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = COLOR_VERT_FORET
4 itmAvant.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = COLOR_VERT_FORET
4 itmAvant.ListSubItems(I_COL_SOUM_MONTAGE).ForeColor = COLOR_VERT_FORET
4 itmAvant.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = COLOR_VERT_FORET
4 itmAvant.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = COLOR_VERT_FORET
4 itmAvant.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = COLOR_VERT_FORET
4 itmAvant.ListSubItems(I_COL_SOUM_TEMPS).ForeColor = COLOR_VERT_FORET
4  itmAvant.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = COLOR_VERT_FORET
4  itmAvant.ListSubItems(I_COL_SOUM_DATE_COMMANDE).ForeColor = COLOR_VERT_FORET
4  itmAvant.ListSubItems(I_COL_SOUM_DATE_REQUISE).ForeColor = COLOR_VERT_FORET
4  itmAvant.ListSubItems(I_COL_SOUM_NOM_COMMANDE).ForeColor = COLOR_VERT_FORET
4  itmAvant.ListSubItems(I_COL_SOUM_NO_SEQUENTIEL).ForeColor = COLOR_VERT_FORET

4  itmAvant.ListSubItems(I_COL_SOUM_PIECE).Bold = True
4  itmAvant.ListSubItems(I_COL_SOUM_DESCR).Bold = True
4  itmAvant.ListSubItems(I_COL_SOUM_DISTRIB).Bold = True
50 itmAvant.ListSubItems(I_COL_SOUM_ESCOMPTE).Bold = True
50 itmAvant.ListSubItems(I_COL_SOUM_MANUFACT).Bold = True
 itmAvant.ListSubItems(I_COL_SOUM_MONTAGE).Bold = True
 itmAvant.ListSubItems(I_COL_SOUM_PRIX_LIST).Bold = True
 itmAvant.ListSubItems(I_COL_SOUM_PRIX_NET).Bold = True
 itmAvant.ListSubItems(I_COL_SOUM_PROFIT).Bold = True
 itmAvant.ListSubItems(I_COL_SOUM_TEMPS).Bold = True
 itmAvant.ListSubItems(I_COL_SOUM_TOTAL).Bold = True
 itmAvant.ListSubItems(I_COL_SOUM_DATE_COMMANDE).Bold = True
 itmAvant.ListSubItems(I_COL_SOUM_DATE_REQUISE).Bold = True
 itmAvant.ListSubItems(I_COL_SOUM_NOM_COMMANDE).Bold = True
 itmAvant.ListSubItems(I_COL_SOUM_NO_SEQUENTIEL).Bold = True

5  Call lvwSoumission.Refresh

5  Call CalculerPrix

5  Exit Sub

Oups:

5  wOups "frmProjSoumElec", "AnnulerCommande", Err, Err.number, Err.Description
End Sub

Private Sub cmdEffacerForfait_Click()

 On Error GoTo Oups

 txtForfait.Text = ""
 lblForfaitInitiale.Caption = ""

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "cmdEffacerForfait_Click", Err, Err.number, Err.Description
End Sub

Private Sub CopierPiece()

 On Error GoTo Oups

 Dim itmCopier As ListItem
 Dim iNbreSelect As Integer
 Dim iCompteur As Integer
 Dim iNbreSelected As Integer
 Dim iIndex As Integer

 For iCompteur = 1 To lvwSoumission.ListItems.count
 If lvwSoumission.ListItems(iCompteur).Selected = True Then
 If lvwSoumission.ListItems(iCompteur).Tag = "" Or lvwSoumission.ListItems(iCompteur).SubItems(I_COL_SOUM_PIECE) = "" Then
 Call MsgBox("Impossible de copier, la sélection contient une section ou une sous-section!", vbOKOnly, "Erreur")

 Exit Sub
  Else
  iNbreSelected = iNbreSelected + 1
  End If
  End If
  Next

  Screen.MousePointer = vbHourglass

  m_iNbreCopie = iNbreSelected

  ReDim m_arr_tyCopie(0 To iNbreSelected - 1)

10 For iCompteur = 1 To lvwSoumission.ListItems.count
1 If lvwSoumission.ListItems(iCompteur).Selected = True Then
 Set itmCopier = lvwSoumission.ListItems(iCompteur)

 m_arr_tyCopie(iIndex).lColor = itmCopier.ListSubItems(I_COL_SOUM_PIECE).ForeColor

 m_arr_tyCopie(iIndex).bChecked = itmCopier.Checked

 m_arr_tyCopie(iIndex).sQuantite = itmCopier.Text

 m_arr_tyCopie(iIndex).sPiece = itmCopier.SubItems(I_COL_SOUM_PIECE)

 m_arr_tyCopie(iIndex).sDescr = itmCopier.SubItems(I_COL_SOUM_DESCR)
 m_arr_tyCopie(iIndex).sDescrTag = itmCopier.ListSubItems(I_COL_SOUM_DESCR).Tag

 m_arr_tyCopie(iIndex).sManufact = itmCopier.SubItems(I_COL_SOUM_MANUFACT)
 
 m_arr_tyCopie(iIndex).sPrixList = itmCopier.SubItems(I_COL_SOUM_PRIX_LIST)
 m_arr_tyCopie(iIndex).sPrixListTag = itmCopier.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag

 m_arr_tyCopie(iIndex).sEscompte = itmCopier.SubItems(I_COL_SOUM_ESCOMPTE)

 m_arr_tyCopie(iIndex).sPrixNet = itmCopier.SubItems(I_COL_SOUM_PRIX_NET)

 m_arr_tyCopie(iIndex).sFRS = itmCopier.SubItems(I_COL_SOUM_DISTRIB)
 m_arr_tyCopie(iIndex).sFRSTag = itmCopier.ListSubItems(I_COL_SOUM_DISTRIB).Tag

 m_arr_tyCopie(iIndex).sTemps = itmCopier.SubItems(I_COL_SOUM_TEMPS)

 m_arr_tyCopie(iIndex).sMontage = itmCopier.SubItems(I_COL_SOUM_MONTAGE)

 m_arr_tyCopie(iIndex).sTotal = itmCopier.SubItems(I_COL_SOUM_TOTAL)

1  m_arr_tyCopie(iIndex).sProfit = itmCopier.SubItems(I_COL_SOUM_PROFIT)

 iIndex = iIndex + 1
 End If
Next

Screen.MousePointer = vbDefault

Exit Sub

Oups:

wOups "frmProjSoumElec", "CopierPiece", Err, Err.number, Err.Description
End Sub

Private Sub CollerPiece()

 On Error GoTo Oups

 Dim sIDSection As String
 Dim sOrdreSection As String
 Dim sSousSection As String
 Dim itmColler As ListItem
 Dim iCompteur As Integer
 Dim iIndexSelected As Integer
 Dim iIndex As Integer

 'Pour savoir s'il y a quelque chose à coller
 If m_iNbreCopie = 0 Then
 Exit Sub
 End If
 
  iIndexSelected = lvwSoumission.SelectedItem.Index
 
  If iIndexSelected >= 3 Then
  If lvwSoumission.SelectedItem.Tag = vbNullString Then
  iIndex = iIndexSelected - 1
  Else
  If lvwSoumission.SelectedItem.SubItems(I_COL_SOUM_PIECE) = vbNullString Then
  If lvwSoumission.ListItems(iIndexSelected - 1).Tag = "" Then
  Call MsgBox("Impossible de coller la pièce entre une section et une sous-section!", vbOKOnly, "Erreur")

 Exit Sub
 Else
 iIndex = iIndexSelected - 1
 End If
 Else
 iIndex = iIndexSelected
 End If
 End If
Else
 Call MsgBox("Emplacement incorrect!", vbOKOnly, "Erreur")

 Exit Sub
End If

1  sIDSection = lvwSoumission.ListItems(iIndex).Tag
sOrdreSection = lvwSoumission.ListItems(iIndex).ListSubItems(I_COL_SOUM_MANUFACT).Tag
 sSousSection = lvwSoumission.ListItems(iIndex).ListSubItems(I_COL_SOUM_PIECE).Tag

Screen.MousePointer = vbHourglass

 For iCompteur = 0 To UBound(m_arr_tyCopie)
 Set itmColler = lvwSoumission.ListItems.Add(iIndexSelected + iCompteur)

 itmColler.Checked = m_arr_tyCopie(iCompteur).bChecked

1  itmColler.Text = m_arr_tyCopie(iCompteur).sQuantite
 itmColler.Tag = sIDSection

 itmColler.SubItems(I_COL_SOUM_PIECE) = m_arr_tyCopie(iCompteur).sPiece
 itmColler.ListSubItems(I_COL_SOUM_PIECE).Tag = sSousSection

 itmColler.SubItems(I_COL_SOUM_DESCR) = m_arr_tyCopie(iCompteur).sDescr
 itmColler.ListSubItems(I_COL_SOUM_DESCR).Tag = m_arr_tyCopie(iCompteur).sDescrTag

 itmColler.SubItems(I_COL_SOUM_MANUFACT) = m_arr_tyCopie(iCompteur).sManufact
 itmColler.ListSubItems(I_COL_SOUM_MANUFACT).Tag = sOrdreSection
 
 itmColler.SubItems(I_COL_SOUM_PRIX_LIST) = m_arr_tyCopie(iCompteur).sPrixList
 itmColler.ListSubItems(I_COL_SOUM_PRIX_LIST).Tag = m_arr_tyCopie(iCompteur).sPrixListTag

 itmColler.SubItems(I_COL_SOUM_ESCOMPTE) = m_arr_tyCopie(iCompteur).sEscompte

 itmColler.SubItems(I_COL_SOUM_PRIX_NET) = m_arr_tyCopie(iCompteur).sPrixNet

 itmColler.SubItems(I_COL_SOUM_DISTRIB) = m_arr_tyCopie(iCompteur).sFRS
itmColler.ListSubItems(I_COL_SOUM_DISTRIB).Tag = m_arr_tyCopie(iCompteur).sFRSTag

 itmColler.SubItems(I_COL_SOUM_TEMPS) = m_arr_tyCopie(iCompteur).sTemps

itmColler.SubItems(I_COL_SOUM_MONTAGE) = m_arr_tyCopie(iCompteur).sMontage

 itmColler.SubItems(I_COL_SOUM_TOTAL) = m_arr_tyCopie(iCompteur).sTotal

itmColler.SubItems(I_COL_SOUM_PROFIT) = m_arr_tyCopie(iCompteur).sProfit

 If m_eType = TYPE_PROJET Then
 itmColler.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = m_arr_tyCopie(iCompteur).lColor
 itmColler.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = m_arr_tyCopie(iCompteur).lColor
 itmColler.ListSubItems(I_COL_SOUM_MONTAGE).ForeColor = m_arr_tyCopie(iCompteur).lColor
itmColler.ListSubItems(I_COL_SOUM_TEMPS).ForeColor = m_arr_tyCopie(iCompteur).lColor
 itmColler.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = m_arr_tyCopie(iCompteur).lColor
 itmColler.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = m_arr_tyCopie(iCompteur).lColor
 itmColler.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = m_arr_tyCopie(iCompteur).lColor
 itmColler.ListSubItems(I_COL_SOUM_DESCR).ForeColor = m_arr_tyCopie(iCompteur).lColor
 itmColler.ListSubItems(I_COL_SOUM_PIECE).ForeColor = m_arr_tyCopie(iCompteur).lColor
 itmColler.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = m_arr_tyCopie(iCompteur).lColor
 itmColler.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = m_arr_tyCopie(iCompteur).lColor
 Else
 itmColler.ListSubItems(I_COL_SOUM_PROFIT).ForeColor = COLOR_NOIR
 itmColler.ListSubItems(I_COL_SOUM_TOTAL).ForeColor = COLOR_NOIR
 itmColler.ListSubItems(I_COL_SOUM_MONTAGE).ForeColor = COLOR_NOIR
 itmColler.ListSubItems(I_COL_SOUM_TEMPS).ForeColor = COLOR_NOIR
 itmColler.ListSubItems(I_COL_SOUM_DISTRIB).ForeColor = COLOR_NOIR
 itmColler.ListSubItems(I_COL_SOUM_PRIX_NET).ForeColor = COLOR_NOIR
 itmColler.ListSubItems(I_COL_SOUM_ESCOMPTE).ForeColor = COLOR_NOIR
 itmColler.ListSubItems(I_COL_SOUM_DESCR).ForeColor = COLOR_NOIR
 itmColler.ListSubItems(I_COL_SOUM_PIECE).ForeColor = COLOR_NOIR
 itmColler.ListSubItems(I_COL_SOUM_MANUFACT).ForeColor = COLOR_NOIR
 itmColler.ListSubItems(I_COL_SOUM_PRIX_LIST).ForeColor = COLOR_NOIR
4 End If

4 Call lvwSoumission.Refresh
4 Next

4 Call CalculerTempsFabrication

4 Call CalculerPrix

4 Screen.MousePointer = vbDefault

4 Exit Sub

Oups:

4 wOups "frmProjSoumElec", "CollerPiece", Err, Err.number, Err.Description
End Sub

Private Sub Deselect()
 
 On Error GoTo Oups
 
 Dim iCompteur As Integer
 
 If lvwSoumission.ListItems.count > 0 Then
 For iCompteur = 1 To lvwSoumission.ListItems.count
 lvwSoumission.ListItems(iCompteur).Selected = False
 Next
 End If

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "Deselect", Err, Err.number, Err.Description
End Sub

Private Sub txtPrixSpecial_Change()

 On Error GoTo Oups
 'Quand le contenu du prix spécial change
 
 'Si la longueur du texte écrit est plus grand que 0
 If Len(txtPrixSpecial.Text) > 0 Then
 'On vide l'escompte, le prix net et on les désactive
 mskEscompte.Text = vbNullString
 txtPrixNet.Text = vbNullString
 
 mskEscompte.Enabled = False
 txtPrixNet.Enabled = False
 Else
 'Sinon, on active escompte et prix net
 mskEscompte.Enabled = True
 txtPrixNet.Enabled = True
 End If

 Exit Sub

Oups:

  wOups "frmProjSoumElec", "txtPrixSpecial_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtPrixSpecial_LostFocus()

 On Error GoTo Oups

 txtPrixSpecial.Text = Replace(txtPrixSpecial.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmProjSoumElec", "txtPrixSpecial_LostFocus", Err, Err.number, Err.Description
End Sub

Private Function ValiderFormatElectrique(ByVal sNoProjSoum As String) As Boolean
 
 On Error GoTo Oups

 If UCase(Left$(sNoProjSoum, 1)) = "E" Then
 ValiderFormatElectrique = True
 Else
 Call MsgBox("Un numéro électrique doit absolument commencé par 'E' !", vbOKOnly, "Erreur")

 ValiderFormatElectrique = False
 End If

 Exit Function

Oups:

 wOups "FrmProjSoumElec", "ValiderFormatElectrique", Err, Err.number, Err.Description
End Function

Private Function ValiderFormatSoumission(ByVal sNoSoumission As String) As Boolean
 
 On Error GoTo Oups

 If Mid$(sNoSoumission, 3, 1) = "1" Then
 ValiderFormatSoumission = True
 Else
 Call MsgBox("Une soumission doit absolument avoir un '1' comme 3e caractère !", vbOKOnly, "Erreur")

 ValiderFormatSoumission = False
 End If

 Exit Function

Oups:

 wOups "FrmProjSoumElec", "ValiderFormatSoumission", Err, Err.number, Err.Description
End Function

Private Function ValiderFormatJobSansSoum(ByVal sNoProjet As String) As Boolean
 
 On Error GoTo Oups

 If Mid$(sNoProjet, 3, 1) <> "3" And Mid$(sNoProjet, 3, 1) <> "1" Then
 ValiderFormatJobSansSoum = True
 Else
 Call MsgBox("Un projet créé sans soumission ne peut pas être un '" & Mid$(sNoProjet, 2, 2) & "' !", vbOKOnly, "Erreur")

 ValiderFormatJobSansSoum = False
 End If

 Exit Function

Oups:

 wOups "FrmProjSoumElec", "ValiderFormatJobSansSoum", Err, Err.number, Err.Description
End Function

Private Function ValiderFormatJobAvecSoum(ByVal sNoProjet As String) As Boolean
 
 On Error GoTo Oups

 If Mid$(sNoProjet, 3, 1) = "3" Then
 ValiderFormatJobAvecSoum = True
 Else
 Call MsgBox("Un projet créé à partir d'une soumission doit absolument avec un '3' comme 3e caractère!", vbOKOnly, "Errreu")

 ValiderFormatJobAvecSoum = False
 End If

 Exit Function

Oups:

 wOups "FrmProjSoumElec", "ValiderFormatJobAvecSoum", Err, Err.number, Err.Description
End Function

Private Function ValiderFormatJobExtra(ByVal sNoProjet As String) As Boolean
 
 On Error GoTo Oups

 If CInt(Right$(sNoProjet, 2)) >= 50 And CInt(Right$(sNoProjet, 2)) <=   Then
 ValiderFormatJobExtra = True
 Else
 Call MsgBox("L'entension d'un extra doit être compris entre 50 et    !", vbOKOnly, "Erreur")

 ValiderFormatJobExtra = False
 End If

 Exit Function

Oups:

 wOups "FrmProjSoumElec", "ValiderFormatJobExtra", Err, Err.number, Err.Description
End Function

Private Sub AjouterProjetAuCumulatif()

 On Error GoTo Oups

 Dim sNoCumulatif As String
 Dim rstProj As ADODB.Recordset
 Dim rstPieces As ADODB.Recordset
 Dim rstProjCumulatif As ADODB.Recordset
 Dim rstPiecesCumulatif As ADODB.Recordset
 Dim rstProjSoum As ADODB.Recordset
 Dim rstEmploye As ADODB.Recordset
 Dim rstSoum As ADODB.Recordset
 Dim bCumulatifExiste As Boolean
 Dim dblNbreManuel As Double
  Dim dblPrixEmballage As Double
  Dim dblTotalManuel As Double
  Dim dblForfait As Double

  sNoCumulatif = Left$(txtNoProjSoum.Text, 7) & "99"

  Set rstProj = New ADODB.Recordset
  Set rstPieces = New ADODB.Recordset
  Set rstProjCumulatif = New ADODB.Recordset
  Set rstPiecesCumulatif = New ADODB.Recordset

10 Call rstProjCumulatif.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = '" & sNoCumulatif & "'", g_connData, adOpenDynamic, adLockOptimistic)

If rstProjCumulatif.EOF Then
 bCumulatifExiste = False

 Call rstProjCumulatif.AddNew

 rstProjCumulatif.Fields("IDProjet") = sNoCumulatif

 'Ouverture du projet -01 pour voir la soumission reliée pour ensuite assigner
 'la soumission -9  avec le projet -99
 Call rstProj.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = '" & Left$(txtNoProjSoum.Text, 6) & "-01'", g_connData, adOpenForwardOnly, adLockReadOnly)

 If Not rstProj.EOF Then
 If Not IsNull(rstProj.Fields("IDSoumission")) Then
 If Len(rstProj.Fields("IDSoumission")) >=   Then
 Set rstSoum = New ADODB.Recordset

 Call rstSoum.Open("SELECT * FROM GrbSoumissionElec WHERE IDSoumission = '" & Left$(rstProj.Fields("IDSoumission"), 6) & "-99'", g_connData, adOpenForwardOnly, adLockReadOnly)

 If Not rstSoum.EOF Then
 rstProjCumulatif.Fields("IDSoumission") = rstSoum.Fields("IDSoumission")
 End If

 Call rstSoum.Close
 Set rstSoum = Nothing
 End If
 End If
 End If

1  Call rstProj.Close

 Call rstProj.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = '" & txtNoProjSoum.Text & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

 rstProjCumulatif.Fields("IDClient") = rstProj.Fields("IDClient")
 rstProjCumulatif.Fields("IDContact") = rstProj.Fields("IDContact")

 rstProjCumulatif.Fields("TauxDessin") = rstProj.Fields("TauxDessin")
 rstProjCumulatif.Fields("TauxFabrication") = rstProj.Fields("TauxFabrication")
 rstProjCumulatif.Fields("TauxAssemblage") = rstProj.Fields("TauxAssemblage")
 rstProjCumulatif.Fields("TauxProgInterface") = rstProj.Fields("TauxProgInterface")
 rstProjCumulatif.Fields("TauxProgAutomate") = rstProj.Fields("TauxProgAutomate")
 rstProjCumulatif.Fields("TauxProgRobot") = rstProj.Fields("TauxProgRobot")
 rstProjCumulatif.Fields("TauxVision") = rstProj.Fields("TauxVision")
 rstProjCumulatif.Fields("TauxTest") = rstProj.Fields("TauxTest")
 rstProjCumulatif.Fields("TauxInstallation") = rstProj.Fields("TauxInstallation")
rstProjCumulatif.Fields("TauxMiseService") = rstProj.Fields("TauxMiseService")
 rstProjCumulatif.Fields("TauxFormation") = rstProj.Fields("TauxFormation")
rstProjCumulatif.Fields("TauxGestion") = rstProj.Fields("TauxGestion")
 rstProjCumulatif.Fields("TauxShipping") = rstProj.Fields("TauxShipping")

rstProjCumulatif.Fields("Transport") = rstProj.Fields("Transport")

 rstProjCumulatif.Fields("Profit") = rstProj.Fields("Profit")
rstProjCumulatif.Fields("imprevue") = rstProj.Fields("imprevue")
 rstProjCumulatif.Fields("commission") = rstProj.Fields("commission")

Call rstProj.Close

3 rstProjCumulatif.Fields("Description") = "Cumulatif de " & Left$(txtNoProjSoum.Text, 6)

 Set rstEmploye = New ADODB.Recordset

 Call rstEmploye.Open("SELECT * FROM Grbemployés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

 rstProjCumulatif.Fields("creer") = ConvertDate(Date)

 rstProjCumulatif.Fields("creer_par") = rstEmploye.Fields("NoEmploye")

 Call rstEmploye.Close
 Set rstEmploye = Nothing

 Call rstProjCumulatif.Update

 Set rstProjSoum = New ADODB.Recordset

 Call rstProjSoum.Open("SELECT * FROM GrbProjSoum", g_connData, adOpenDynamic, adLockOptimistic)

 Call rstProjSoum.AddNew

rstProjSoum.Fields("IDProjSoum") = sNoCumulatif
 rstProjSoum.Fields("NoClient") = rstProjCumulatif.Fields("IDClient")
rstProjSoum.Fields("Description") = rstProjCumulatif.Fields("Description")
 rstProjSoum.Fields("DateOuverture") = ConvertDate(Date)
rstProjSoum.Fields("Ouvert") = True
 rstProjSoum.Fields("Verrouillé") = True
 rstProjSoum.Fields("Type") = "P"

 Call rstProjSoum.Update
 
Call rstProjSoum.Close
4 Set rstProjSoum = Nothing
4 Else
4 bCumulatifExiste = True
4 End If

4 rstProj.CursorLocation = adUseClient

4 Call rstProj.Open("SELECT * FROM GrbProjetElec WHERE LEFT(IDProjet, 6) = '" & Left$(txtNoProjSoum.Text, 6) & "' AND RIGHT(IDProjet, 2) <> '99'", g_connData, adOpenForwardOnly, adLockReadOnly)

4 If rstProj.RecordCount = 1 Then
4 rstProjCumulatif.Fields("NbreManuel") = rstProj.Fields("NbreManuel")

4 rstProjCumulatif.Fields("PrixEmballage") = rstProj.Fields("PrixEmballage")

4 rstProjCumulatif.Fields("total_manuel") = rstProj.Fields("total_manuel")

4 rstProjCumulatif.Fields("MontantForfait") = rstProj.Fields("MontantForfait")
4  Else
4  Do While Not rstProj.EOF
4  If Not IsNull(rstProj.Fields("NbreManuel")) Then
4  dblNbreManuel = dblNbreManuel + CDbl(rstProj.Fields("NbreManuel"))
4  End If

4  If Not IsNull(rstProj.Fields("PrixEmballage")) Then
4  dblPrixEmballage = dblPrixEmballage + CDbl(rstProj.Fields("PrixEmballage"))
4  End If

50 If Not IsNull(rstProj.Fields("total_manuel")) Then
 dblTotalManuel = dblTotalManuel + CDbl(rstProj.Fields("total_manuel"))
 End If

 If Not IsNull(rstProj.Fields("MontantForfait")) Then
 If IsNumeric(rstProj.Fields("MontantForfait")) Then
 dblForfait = dblForfait + CDbl(rstProj.Fields("MontantForfait"))
 End If
 End If

 Call rstProj.MoveNext
 Loop

 rstProjCumulatif.Fields("NbreManuel") = dblNbreManuel
 rstProjCumulatif.Fields("PrixEmballage") = dblPrixEmballage
5  rstProjCumulatif.Fields("total_manuel") = dblTotalManuel
5  rstProjCumulatif.Fields("MontantForfait") = dblForfait
5  End If

5  Call rstProj.Close

5  Call rstProjCumulatif.Update

5  Call rstProjCumulatif.Close

 'AJOUT DES PIÈCES
5  Call rstPiecesCumulatif.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & sNoCumulatif & "'", g_connData, adOpenDynamic, adLockOptimistic)

5  If bCumulatifExiste = True Then
60 Call g_connData.Execute("DELETE * FROM GrbProjet_Pieces WHERE IDProjet = '" & sNoCumulatif & "' AND Provenance = '" & Right$(txtNoProjSoum.Text, 2) & "'")

  Call rstPieces.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & txtNoProjSoum.Text & "' AND Provenance Is Null OR Provenance = '' ORDER BY NuméroLigne", g_connData, adOpenForwardOnly, adLockReadOnly)
  Else
  Call rstPieces.Open("SELECT * FROM GrbProjet_Pieces WHERE LEFT(IDProjet, 6) = '" & Left$(txtNoProjSoum.Text, 6) & "' AND RIGHT(IDProjet, 2) <> '99' AND Provenance Is Null OR Provenance = '' ORDER BY CInt(OrdreSection), NuméroLigne", g_connData, adOpenForwardOnly, adLockReadOnly)
  End If

  Do While Not rstPieces.EOF
  Call rstPiecesCumulatif.AddNew

  rstPiecesCumulatif.Fields("IDProjet") = sNoCumulatif
  rstPiecesCumulatif.Fields("IDSection") = rstPieces.Fields("IDSection")
  rstPiecesCumulatif.Fields("NumItem") = rstPieces.Fields("NumItem")
  rstPiecesCumulatif.Fields("Qté") = rstPieces.Fields("Qté")
  rstPiecesCumulatif.Fields("Desc_FR") = rstPieces.Fields("Desc_FR")
6  rstPiecesCumulatif.Fields("Desc_EN") = rstPieces.Fields("Desc_EN")
6  rstPiecesCumulatif.Fields("Manufact") = rstPieces.Fields("Manufact")
6  rstPiecesCumulatif.Fields("Prix_list") = rstPieces.Fields("Prix_list")
6  rstPiecesCumulatif.Fields("Escompte") = rstPieces.Fields("Escompte")
6  rstPiecesCumulatif.Fields("Prix_net") = rstPieces.Fields("Prix_net")
6  rstPiecesCumulatif.Fields("IDFRS") = rstPieces.Fields("IDFRS")
6  rstPiecesCumulatif.Fields("Temps") = rstPieces.Fields("Temps")
6  rstPiecesCumulatif.Fields("Temps_Total") = rstPieces.Fields("Temps_Total")
70 rstPiecesCumulatif.Fields("Prix_total") = rstPieces.Fields("Prix_total")
  rstPiecesCumulatif.Fields("Profit_Argent") = rstPieces.Fields("Profit_Argent")
  rstPiecesCumulatif.Fields("SousSection") = rstPieces.Fields("SousSection")
  rstPiecesCumulatif.Fields("OrdreSection") = rstPieces.Fields("OrdreSection")
  rstPiecesCumulatif.Fields("NuméroLigne") = rstPieces.Fields("NuméroLigne")
  rstPiecesCumulatif.Fields("PrixOrigine") = rstPieces.Fields("PrixOrigine")
  rstPiecesCumulatif.Fields("Type") = rstPieces.Fields("Type")
  rstPiecesCumulatif.Fields("Visible") = rstPieces.Fields("Visible")
  rstPiecesCumulatif.Fields("Commentaire") = rstPieces.Fields("Commentaire")
  rstPiecesCumulatif.Fields("Devise") = rstPieces.Fields("Devise")
  rstPiecesCumulatif.Fields("Provenance") = Right(rstPieces.Fields("IDProjet"), 2)

  Call rstPiecesCumulatif.Update

   Call rstPieces.MoveNext
   Loop

7  Call rstPiecesCumulatif.Close
7  Call rstPieces.Close

7  Set rstProj = Nothing
7  Set rstPieces = Nothing
7  Set rstProjCumulatif = Nothing
7  Set rstPiecesCumulatif = Nothing

80 Call CalculerTotalRecordset(sNoCumulatif)

80 If bCumulatifExiste = False Then
  If cmbOuvertFerme.ListIndex = I_CMB_OUVERT Then
  Call RemplirComboProjSoum(txtNoProjSoum.Text)
  End If
  End If

  Exit Sub

Oups:

  wOups "FrmProjSoumElec", "AjouterProjetAuCumulatif", Err, Err.number, Err.Description
End Sub

Private Sub AjouterSoumissionAuCumulatif()

 On Error GoTo Oups

 Dim sNoCumulatif As String
 Dim rstSoum As ADODB.Recordset
 Dim rstPieces As ADODB.Recordset
 Dim rstSoumCumulatif As ADODB.Recordset
 Dim rstPiecesCumulatif As ADODB.Recordset
 Dim rstProjSoum As ADODB.Recordset
 Dim rstEmploye As ADODB.Recordset
 Dim bCumulatifExiste As Boolean
 Dim dblNbreManuel As Double
 Dim dblTempsDessin As Double
  Dim dblTempsFabrication As Double
  Dim dblTempsAssemblage As Double
  Dim dblTempsProgInterface As Double
  Dim dblTempsProgAutomate As Double
  Dim dblTempsProgRobot As Double
  Dim dblTempsVision As Double
  Dim dblTempsTest As Double
  Dim dblTempsInstallation As Double
10 Dim dblTempsMiseService As Double
Dim dblTempsFormation As Double
Dim dblTempsGestion As Double
Dim dblTempsShipping As Double
Dim dblTempsTransport As Double
Dim dblTempsUniteMobile As Double
Dim dblTotalHebergement As Double
Dim dblTotalRepas As Double
Dim dblPrixEmballage As Double
Dim dblTotalManuel As Double
Dim dblForfait As Double

sNoCumulatif = Left$(txtNoProjSoum.Text, 7) & "99"

1  Set rstSoum = New ADODB.Recordset
Set rstPieces = New ADODB.Recordset
 Set rstSoumCumulatif = New ADODB.Recordset
Set rstPiecesCumulatif = New ADODB.Recordset

 Call rstSoumCumulatif.Open("SELECT * FROM GrbSoumissionElec WHERE IDSoumission = '" & sNoCumulatif & "'", g_connData, adOpenDynamic, adLockOptimistic)

If rstSoumCumulatif.EOF Then
 bCumulatifExiste = False

1  Call rstSoumCumulatif.AddNew

 rstSoumCumulatif.Fields("IDSoumission") = sNoCumulatif

 Call rstSoum.Open("SELECT * FROM GrbSoumissionElec WHERE IDSoumission = '" & txtNoProjSoum.Text & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

 rstSoumCumulatif.Fields("IDClient") = rstSoum.Fields("IDClient")
 rstSoumCumulatif.Fields("IDContact") = rstSoum.Fields("IDContact")

 rstSoumCumulatif.Fields("TauxDessin") = rstSoum.Fields("TauxDessin")
 rstSoumCumulatif.Fields("TauxFabrication") = rstSoum.Fields("TauxFabrication")
 rstSoumCumulatif.Fields("TauxAssemblage") = rstSoum.Fields("TauxAssemblage")
 rstSoumCumulatif.Fields("TauxProgInterface") = rstSoum.Fields("TauxProgInterface")
 rstSoumCumulatif.Fields("TauxProgAutomate") = rstSoum.Fields("TauxProgAutomate")
 rstSoumCumulatif.Fields("TauxProgRobot") = rstSoum.Fields("TauxProgRobot")
 rstSoumCumulatif.Fields("TauxVision") = rstSoum.Fields("TauxVision")
 rstSoumCumulatif.Fields("TauxTest") = rstSoum.Fields("TauxTest")
rstSoumCumulatif.Fields("TauxInstallation") = rstSoum.Fields("TauxInstallation")
 rstSoumCumulatif.Fields("TauxMiseService") = rstSoum.Fields("TauxMiseService")
rstSoumCumulatif.Fields("TauxFormation") = rstSoum.Fields("TauxFormation")
 rstSoumCumulatif.Fields("TauxGestion") = rstSoum.Fields("TauxGestion")
rstSoumCumulatif.Fields("TauxShipping") = rstSoum.Fields("TauxShipping")

 rstSoumCumulatif.Fields("TauxHebergement1") = rstSoum.Fields("TauxHebergement1")
rstSoumCumulatif.Fields("TauxHebergement2") = rstSoum.Fields("TauxHebergement2")
 rstSoumCumulatif.Fields("TauxRepas") = rstSoum.Fields("TauxRepas")
rstSoumCumulatif.Fields("TauxTransport") = rstSoum.Fields("TauxTransport")
3 rstSoumCumulatif.Fields("TauxUniteMobile") = rstSoum.Fields("TauxUniteMobile")

 rstSoumCumulatif.Fields("Transport") = rstSoum.Fields("Transport")

 rstSoumCumulatif.Fields("Profit") = rstSoum.Fields("Profit")
 rstSoumCumulatif.Fields("imprevue") = rstSoum.Fields("imprevue")
 rstSoumCumulatif.Fields("commission") = rstSoum.Fields("commission")

 Call rstSoum.Close

 rstSoumCumulatif.Fields("Description") = "Cumulatif de " & Left$(txtNoProjSoum.Text, 6)

 Set rstEmploye = New ADODB.Recordset

 Call rstEmploye.Open("SELECT * FROM Grbemployés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

 rstSoumCumulatif.Fields("creer") = ConvertDate(Date)

 rstSoumCumulatif.Fields("creer_par") = rstEmploye.Fields("NoEmploye")

Call rstEmploye.Close
 Set rstEmploye = Nothing

Call rstSoumCumulatif.Update

 Set rstProjSoum = New ADODB.Recordset

Call rstProjSoum.Open("SELECT * FROM GrbProjSoum", g_connData, adOpenDynamic, adLockOptimistic)

 Call rstProjSoum.AddNew

 rstProjSoum.Fields("IDProjSoum") = sNoCumulatif
 rstProjSoum.Fields("NoClient") = rstSoumCumulatif.Fields("IDClient")
rstProjSoum.Fields("Description") = rstSoumCumulatif.Fields("Description")
4 rstProjSoum.Fields("DateOuverture") = ConvertDate(Date)
4 rstProjSoum.Fields("Ouvert") = True
4 rstProjSoum.Fields("Verrouillé") = True
4 rstProjSoum.Fields("Type") = "S"

4 Call rstProjSoum.Update
 
4 Call rstProjSoum.Close
4 Set rstProjSoum = Nothing
4 Else
4 bCumulatifExiste = True
4 End If
 
4 rstSoum.CursorLocation = adUseClient
 
4  Call rstSoum.Open("SELECT * FROM GrbSoumissionElec WHERE LEFT(IDSoumission, 6) = '" & Left$(txtNoProjSoum.Text, 6) & "' AND RIGHT(IDSoumission, 2) <> '99'", g_connData, adOpenForwardOnly, adLockReadOnly)

4  If rstSoum.RecordCount = 1 Then
4  rstSoumCumulatif.Fields("NbreManuel") = rstSoum.Fields("NbreManuel")

4  rstSoumCumulatif.Fields("TempsDessin") = rstSoum.Fields("TempsDessin")

4  If rstSoum.Fields("SansTemps") = False Then
4  rstSoumCumulatif.Fields("TempsFabrication") = rstSoum.Fields("TempsFabrication")
4  Else
4  rstSoumCumulatif.Fields("TempsFabrication") = 0
50 End If

5 rstSoumCumulatif.Fields("TempsAssemblage") = rstSoum.Fields("TempsAssemblage")
 rstSoumCumulatif.Fields("TempsProgInterface") = rstSoum.Fields("TempsProgInterface")
 rstSoumCumulatif.Fields("TempsProgAutomate") = rstSoum.Fields("TempsProgAutomate")
 rstSoumCumulatif.Fields("TempsProgRobot") = rstSoum.Fields("TempsProgRobot")
 rstSoumCumulatif.Fields("TempsVision") = rstSoum.Fields("TempsVision")
 rstSoumCumulatif.Fields("TempsTest") = rstSoum.Fields("TempsTest")
 rstSoumCumulatif.Fields("TempsInstallation") = rstSoum.Fields("TempsInstallation")
 rstSoumCumulatif.Fields("TempsMiseService") = rstSoum.Fields("TempsMiseService")
 rstSoumCumulatif.Fields("TempsFormation") = rstSoum.Fields("TempsFormation")
 rstSoumCumulatif.Fields("TempsGestion") = rstSoum.Fields("TempsGestion")
 rstSoumCumulatif.Fields("TempsShipping") = rstSoum.Fields("TempsShipping")

5  rstSoumCumulatif.Fields("NbrePersonne") = rstSoum.Fields("NbrePersonne")
5  rstSoumCumulatif.Fields("TempsHebergement") = rstSoum.Fields("TempsHebergement")
5  rstSoumCumulatif.Fields("TempsRepas") = rstSoum.Fields("TempsRepas")
5  rstSoumCumulatif.Fields("TempsTransport") = rstSoum.Fields("TempsTransport")
5  rstSoumCumulatif.Fields("TempsUniteMobile") = rstSoum.Fields("TempsUniteMobile")

5  rstSoumCumulatif.Fields("TotalHebergement") = rstSoum.Fields("TotalHebergement")
5  rstSoumCumulatif.Fields("TotalRepas") = rstSoum.Fields("TotalRepas")
5  rstSoumCumulatif.Fields("PrixEmballage") = rstSoum.Fields("PrixEmballage")

60 rstSoumCumulatif.Fields("total_manuel") = rstSoum.Fields("total_manuel")

  rstSoumCumulatif.Fields("MontantForfait") = rstSoum.Fields("MontantForfait")
  Else
  Do While Not rstSoum.EOF
  If Not IsNull(rstSoum.Fields("NbreManuel")) Then
  dblNbreManuel = dblNbreManuel + CDbl(rstSoum.Fields("NbreManuel"))
  End If

  If Not IsNull(rstSoum.Fields("TempsDessin")) Then
  dblTempsDessin = dblTempsDessin + CDbl(rstSoum.Fields("TempsDessin"))
  End If

  If rstSoum.Fields("SansTemps") = False Then
  If Not IsNull(rstSoum.Fields("TempsFabrication")) Then
6  dblTempsFabrication = dblTempsFabrication + CDbl(rstSoum.Fields("TempsFabrication"))
6  End If
6  End If

6  If Not IsNull(rstSoum.Fields("TempsAssemblage")) Then
6  dblTempsAssemblage = dblTempsAssemblage + CDbl(rstSoum.Fields("TempsAssemblage"))
6  End If

6  If Not IsNull(rstSoum.Fields("TempsProgInterface")) Then
6  dblTempsProgInterface = dblTempsProgInterface + CDbl(rstSoum.Fields("TempsProgInterface"))
70 End If

  If Not IsNull(rstSoum.Fields("TempsProgAutomate")) Then
  dblTempsProgAutomate = dblTempsProgAutomate + CDbl(rstSoum.Fields("TempsProgAutomate"))
  End If

  If Not IsNull(rstSoum.Fields("TempsProgRobot")) Then
  dblTempsProgRobot = dblTempsProgRobot + CDbl(rstSoum.Fields("TempsProgRobot"))
  End If

  If Not IsNull(rstSoum.Fields("TempsVision")) Then
  dblTempsVision = dblTempsVision + CDbl(rstSoum.Fields("TempsVision"))
  End If

  If Not IsNull(rstSoum.Fields("TempsTest")) Then
  dblTempsTest = dblTempsTest + CDbl(rstSoum.Fields("TempsTest"))
   End If

   If Not IsNull(rstSoum.Fields("TempsInstallation")) Then
7  dblTempsInstallation = dblTempsInstallation + CDbl(rstSoum.Fields("TempsInstallation"))
7  End If

7  If Not IsNull(rstSoum.Fields("TempsMiseService")) Then
7  dblTempsMiseService = dblTempsMiseService + CDbl(rstSoum.Fields("TempsMiseService"))
7  End If

7  If Not IsNull(rstSoum.Fields("TempsFormation")) Then
80 dblTempsFormation = dblTempsFormation + CDbl(rstSoum.Fields("TempsFormation"))
  End If

  If Not IsNull(rstSoum.Fields("TempsGestion")) Then
  dblTempsGestion = dblTempsGestion + CDbl(rstSoum.Fields("TempsGestion"))
  End If

  If Not IsNull(rstSoum.Fields("TempsShipping")) Then
  dblTempsShipping = dblTempsShipping + CDbl(rstSoum.Fields("TempsShipping"))
  End If

  If Not IsNull(rstSoum.Fields("TempsTransport")) Then
  dblTempsTransport = dblTempsTransport + CDbl(rstSoum.Fields("TempsTransport"))
  End If

  If Not IsNull(rstSoum.Fields("TempsUniteMobile")) Then
   dblTempsUniteMobile = dblTempsUniteMobile + CDbl(rstSoum.Fields("TempsUniteMobile"))
   End If

   If Not IsNull(rstSoum.Fields("TotalHebergement")) Then
   dblTotalHebergement = dblTotalHebergement + CDbl(rstSoum.Fields("TotalHebergement"))
8  End If

8  If Not IsNull(rstSoum.Fields("TotalRepas")) Then
8  dblTotalRepas = dblTotalRepas + CDbl(rstSoum.Fields("TotalRepas"))
8  End If

90 If Not IsNull(rstSoum.Fields("PrixEmballage")) Then
  dblPrixEmballage = dblPrixEmballage + CDbl(rstSoum.Fields("PrixEmballage"))
  End If

  If Not IsNull(rstSoum.Fields("total_manuel")) Then
  dblTotalManuel = dblTotalManuel + CDbl(rstSoum.Fields("total_manuel"))
  End If

  If Not IsNull(rstSoum.Fields("MontantForfait")) Then
  If IsNumeric(rstSoum.Fields("MontantForfait")) Then
  dblForfait = dblForfait + CDbl(rstSoum.Fields("MontantForfait"))
  End If
  End If

  Call rstSoum.MoveNext
 Loop

   rstSoumCumulatif.Fields("NbreManuel") = dblNbreManuel

 rstSoumCumulatif.Fields("TempsDessin") = dblTempsDessin
   rstSoumCumulatif.Fields("TempsFabrication") = dblTempsFabrication
 rstSoumCumulatif.Fields("TempsAssemblage") = dblTempsAssemblage
   rstSoumCumulatif.Fields("TempsProgInterface") = dblTempsProgInterface
 rstSoumCumulatif.Fields("TempsProgAutomate") = dblTempsProgAutomate
9  rstSoumCumulatif.Fields("TempsProgRobot") = dblTempsProgRobot
 rstSoumCumulatif.Fields("TempsVision") = dblTempsVision
10 rstSoumCumulatif.Fields("TempsTest") = dblTempsTest
1 rstSoumCumulatif.Fields("TempsInstallation") = dblTempsInstallation
1 rstSoumCumulatif.Fields("TempsMiseService") = dblTempsMiseService
 rstSoumCumulatif.Fields("TempsFormation") = dblTempsFormation
1 rstSoumCumulatif.Fields("TempsGestion") = dblTempsGestion
 rstSoumCumulatif.Fields("TempsShipping") = dblTempsShipping

1 rstSoumCumulatif.Fields("TempsTransport") = dblTempsTransport
 rstSoumCumulatif.Fields("TempsUniteMobile") = dblTempsUniteMobile

1 rstSoumCumulatif.Fields("TotalHebergement") = dblTotalHebergement
 rstSoumCumulatif.Fields("TotalRepas") = dblTotalRepas
1 rstSoumCumulatif.Fields("PrixEmballage") = dblPrixEmballage

10  rstSoumCumulatif.Fields("total_manuel") = dblTotalManuel

10  rstSoumCumulatif.Fields("MontantForfait") = dblForfait
107End If

10  Call rstSoumCumulatif.Update

10  Call rstSoumCumulatif.Close

10  Call rstSoum.Close

 'AJOUT DES PIÈCES
10  Call rstPiecesCumulatif.Open("SELECT * FROM GrbSoumission_Pieces WHERE IDSoumission = '" & sNoCumulatif & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
10  If bCumulatifExiste = True Then
110 Call g_connData.Execute("DELETE * FROM GrbSoumission_Pieces WHERE IDSoumission = '" & sNoCumulatif & "' AND Provenance = '" & Right$(txtNoProjSoum.Text, 2) & "'")

11 Call rstPieces.Open("SELECT * FROM GrbSoumission_Pieces WHERE IDSoumission = '" & txtNoProjSoum.Text & "' ORDER BY NuméroLigne", g_connData, adOpenForwardOnly, adLockReadOnly)
11 Else
1 Call rstPieces.Open("SELECT * FROM GrbSoumission_Pieces WHERE LEFT(IDSoumission, 6) = '" & Left$(txtNoProjSoum.Text, 6) & "' AND RIGHT(IDSoumission, 2) <> '99' ORDER BY CInt(OrdreSection), NuméroLigne", g_connData, adOpenForwardOnly, adLockReadOnly)
11 End If

11 Do While Not rstPieces.EOF
1 Call rstPiecesCumulatif.AddNew

1 rstPiecesCumulatif.Fields("IDSoumission") = sNoCumulatif
1 rstPiecesCumulatif.Fields("IDSection") = rstPieces.Fields("IDSection")
1 rstPiecesCumulatif.Fields("NumItem") = rstPieces.Fields("NumItem")
1 rstPiecesCumulatif.Fields("Qté") = rstPieces.Fields("Qté")
1 rstPiecesCumulatif.Fields("Desc_FR") = rstPieces.Fields("Desc_FR")
11  rstPiecesCumulatif.Fields("Desc_EN") = rstPieces.Fields("Desc_EN")
1 rstPiecesCumulatif.Fields("Manufact") = rstPieces.Fields("Manufact")
 rstPiecesCumulatif.Fields("Prix_list") = rstPieces.Fields("Prix_list")
1 rstPiecesCumulatif.Fields("Escompte") = rstPieces.Fields("Escompte")
 rstPiecesCumulatif.Fields("Prix_net") = rstPieces.Fields("Prix_net")
1 rstPiecesCumulatif.Fields("IDFRS") = rstPieces.Fields("IDFRS")
 rstPiecesCumulatif.Fields("Temps") = rstPieces.Fields("Temps")
11  rstPiecesCumulatif.Fields("Temps_Total") = rstPieces.Fields("Temps_Total")
 rstPiecesCumulatif.Fields("Prix_total") = rstPieces.Fields("Prix_total")
1 rstPiecesCumulatif.Fields("Profit_Argent") = rstPieces.Fields("Profit_Argent")
1 rstPiecesCumulatif.Fields("SousSection") = rstPieces.Fields("SousSection")
1 rstPiecesCumulatif.Fields("OrdreSection") = rstPieces.Fields("OrdreSection")
1 rstPiecesCumulatif.Fields("NuméroLigne") = rstPieces.Fields("NuméroLigne")
1 rstPiecesCumulatif.Fields("PrixOrigine") = rstPieces.Fields("PrixOrigine")
1 rstPiecesCumulatif.Fields("Type") = rstPieces.Fields("Type")
1 rstPiecesCumulatif.Fields("Visible") = rstPieces.Fields("Visible")
1 rstPiecesCumulatif.Fields("Commentaire") = rstPieces.Fields("Commentaire")
1 rstPiecesCumulatif.Fields("Devise") = rstPieces.Fields("Devise")
1 rstPiecesCumulatif.Fields("Provenance") = Right(rstPieces.Fields("IDSoumission"), 2)

1 Call rstPiecesCumulatif.Update

12  Call rstPieces.MoveNext
12  Loop

12  Call rstPiecesCumulatif.Close
12  Call rstPieces.Close

12  Set rstSoum = Nothing
12  Set rstPieces = Nothing
12  Set rstSoumCumulatif = Nothing
12  Set rstPiecesCumulatif = Nothing

130 Call CalculerTotalRecordset(sNoCumulatif)

130 Exit Sub

Oups:

13 wOups "FrmProjSoumElec", "AjouterSoumissionAuCumulatif", Err, Err.number, Err.Description
End Sub

Private Sub RecreerProjetCumulatif()

 On Error GoTo Oups

 Dim sNoCumulatif As String
 Dim rstProj As ADODB.Recordset
 Dim rstPieces As ADODB.Recordset
 Dim rstProjCumulatif As ADODB.Recordset
 Dim rstPiecesCumulatif As ADODB.Recordset
 Dim rstProjSoum As ADODB.Recordset
 Dim dblNbreManuel As Double
 Dim dblPrixEmballage As Double
 Dim dblTotalManuel As Double
 Dim dblForfait As Double
  Dim bSupprimer As Boolean

  sNoCumulatif = Left$(txtNoProjSoum.Text, 7) & "99"

  Set rstProj = New ADODB.Recordset
  Set rstPieces = New ADODB.Recordset
  Set rstProjCumulatif = New ADODB.Recordset
  Set rstPiecesCumulatif = New ADODB.Recordset

  rstProj.CursorLocation = adUseClient

  Call rstProj.Open("SELECT * FROM GrbProjetElec WHERE LEFT(IDProjet, 6) = '" & Left$(txtNoProjSoum.Text, 6) & "' AND RIGHT(IDProjet, 2) <> '99'", g_connData, adOpenForwardOnly, adLockReadOnly)

10 If rstProj.EOF Then
1 Call g_connData.Execute("DELETE * FROM GrbProjSoum WHERE IDProjSoum = '" & sNoCumulatif & "'")

 Call g_connData.Execute("DELETE * FROM GrbProjet_Pieces WHERE IDProjet = '" & sNoCumulatif & "' AND Type = 'E'")

 'Efface le projet
 Call g_connData.Execute("DELETE * FROM GrbProjetElec WHERE IDProjet = '" & sNoCumulatif & "'")

 bSupprimer = True
Else
 Call rstProjCumulatif.Open("SELECT * FROM GrbProjetElec WHERE IDProjet = '" & sNoCumulatif & "'", g_connData, adOpenDynamic, adLockOptimistic)

 If rstProj.RecordCount = 1 Then
 rstProjCumulatif.Fields("NbreManuel") = rstProj.Fields("NbreManuel")

 rstProjCumulatif.Fields("PrixEmballage") = rstProj.Fields("PrixEmballage")

 rstProjCumulatif.Fields("total_manuel") = rstProj.Fields("total_manuel")

 rstProjCumulatif.Fields("MontantForfait") = rstProj.Fields("MontantForfait")
Else
 Do While Not rstProj.EOF
 If Not IsNull(rstProj.Fields("NbreManuel")) Then
 dblNbreManuel = dblNbreManuel + CDbl(rstProj.Fields("NbreManuel"))
 End If

 If Not IsNull(rstProj.Fields("PrixEmballage")) Then
 dblPrixEmballage = dblPrixEmballage + CDbl(rstProj.Fields("PrixEmballage"))
1  End If

 If Not IsNull(rstProj.Fields("total_manuel")) Then
 dblTotalManuel = dblTotalManuel + CDbl(rstProj.Fields("total_manuel"))
 End If

 If Not IsNull(rstProj.Fields("MontantForfait")) Then
 If IsNumeric(rstProj.Fields("MontantForfait")) Then
 dblForfait = dblForfait + CDbl(rstProj.Fields("MontantForfait"))
 End If
 End If

 Call rstProj.MoveNext
 Loop

 rstProjCumulatif.Fields("NbreManuel") = dblNbreManuel
 rstProjCumulatif.Fields("PrixEmballage") = dblPrixEmballage
 rstProjCumulatif.Fields("total_manuel") = dblTotalManuel
 rstProjCumulatif.Fields("MontantForfait") = dblForfait
End If

 Call rstProj.Close

Call rstProjCumulatif.Update

 Call rstProjCumulatif.Close
2  End If

 'AJOUT DES PIÈCES
Call rstPiecesCumulatif.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & sNoCumulatif & "'", g_connData, adOpenDynamic, adLockOptimistic)

30 Call g_connData.Execute("DELETE * FROM GrbProjet_Pieces WHERE IDProjet = '" & sNoCumulatif & "' AND Provenance = '" & Right$(txtNoProjSoum.Text, 2) & "'")

Call rstPieces.Open("SELECT * FROM GrbProjet_Pieces WHERE IDProjet = '" & txtNoProjSoum.Text & "' AND Provenance Is Null OR Provenance = '' ORDER BY NuméroLigne", g_connData, adOpenForwardOnly, adLockReadOnly)

Do While Not rstPieces.EOF
 Call rstPiecesCumulatif.AddNew

 rstPiecesCumulatif.Fields("IDProjet") = sNoCumulatif
 rstPiecesCumulatif.Fields("IDSection") = rstPieces.Fields("IDSection")
 rstPiecesCumulatif.Fields("NumItem") = rstPieces.Fields("NumItem")
 rstPiecesCumulatif.Fields("Qté") = rstPieces.Fields("Qté")
 rstPiecesCumulatif.Fields("Desc_FR") = rstPieces.Fields("Desc_FR")
 rstPiecesCumulatif.Fields("Desc_EN") = rstPieces.Fields("Desc_EN")
 rstPiecesCumulatif.Fields("Manufact") = rstPieces.Fields("Manufact")
 rstPiecesCumulatif.Fields("Prix_list") = rstPieces.Fields("Prix_list")
rstPiecesCumulatif.Fields("Escompte") = rstPieces.Fields("Escompte")
 rstPiecesCumulatif.Fields("Prix_net") = rstPieces.Fields("Prix_net")
rstPiecesCumulatif.Fields("IDFRS") = rstPieces.Fields("IDFRS")
 rstPiecesCumulatif.Fields("Temps") = rstPieces.Fields("Temps")
rstPiecesCumulatif.Fields("Temps_Total") = rstPieces.Fields("Temps_Total")
 rstPiecesCumulatif.Fields("Prix_total") = rstPieces.Fields("Prix_total")
 rstPiecesCumulatif.Fields("Profit_Argent") = rstPieces.Fields("Profit_Argent")
 rstPiecesCumulatif.Fields("SousSection") = rstPieces.Fields("SousSection")
rstPiecesCumulatif.Fields("OrdreSection") = rstPieces.Fields("OrdreSection")
4 rstPiecesCumulatif.Fields("NuméroLigne") = rstPieces.Fields("NuméroLigne")
4 rstPiecesCumulatif.Fields("PrixOrigine") = rstPieces.Fields("PrixOrigine")
4 rstPiecesCumulatif.Fields("Type") = rstPieces.Fields("Type")
4 rstPiecesCumulatif.Fields("Visible") = rstPieces.Fields("Visible")
4 rstPiecesCumulatif.Fields("Commentaire") = rstPieces.Fields("Commentaire")
4 rstPiecesCumulatif.Fields("Devise") = rstPieces.Fields("Devise")
4 rstPiecesCumulatif.Fields("Provenance") = Right(rstPieces.Fields("IDProjet"), 2)

4 Call rstPiecesCumulatif.Update

4 Call rstPieces.MoveNext
4 Loop

4 Call rstPiecesCumulatif.Close
4  Call rstPieces.Close

4  Set rstProj = Nothing
4  Set rstPieces = Nothing
4  Set rstProjCumulatif = Nothing
4  Set rstPiecesCumulatif = Nothing

4  If bSupprimer = False Then
4  Call CalculerTotalRecordset(sNoCumulatif)
4  End If

50 Exit Sub

Oups:

50 wOups "FrmProjSoumElec", "AjouterProjetAuCumulatif", Err, Err.number, Err.Description
End Sub

Private Sub RecreerSoumissionCumulatif()

 On Error GoTo Oups

 Dim sNoCumulatif As String
 Dim rstSoum As ADODB.Recordset
 Dim rstPieces As ADODB.Recordset
 Dim rstSoumCumulatif As ADODB.Recordset
 Dim rstPiecesCumulatif As ADODB.Recordset
 Dim rstProjSoum As ADODB.Recordset
 Dim dblNbreManuel As Double
 Dim dblTempsDessin As Double
 Dim dblTempsFabrication As Double
 Dim dblTempsAssemblage As Double
  Dim dblTempsProgInterface As Double
  Dim dblTempsProgAutomate As Double
  Dim dblTempsProgRobot As Double
  Dim dblTempsVision As Double
  Dim dblTempsTest As Double
  Dim dblTempsInstallation As Double
  Dim dblTempsMiseService As Double
  Dim dblTempsFormation As Double
10 Dim dblTempsGestion As Double
Dim dblTempsShipping As Double
Dim dblTempsTransport As Double
Dim dblTempsUniteMobile As Double
Dim dblTotalHebergement As Double
Dim dblTotalRepas As Double
Dim dblPrixEmballage As Double
Dim dblTotalManuel As Double
Dim dblForfait As Double
Dim bSupprimer As Boolean

sNoCumulatif = Left$(txtNoProjSoum.Text, 7) & "99"

Set rstSoum = New ADODB.Recordset
1  Set rstPieces = New ADODB.Recordset
Set rstSoumCumulatif = New ADODB.Recordset
 Set rstPiecesCumulatif = New ADODB.Recordset
 
rstSoum.CursorLocation = adUseClient
 
 Call rstSoum.Open("SELECT * FROM GrbSoumissionElec WHERE LEFT(IDSoumission, 6) = '" & Left$(txtNoProjSoum.Text, 6) & "' AND RIGHT(IDSoumission, 2) <> '99'", g_connData, adOpenForwardOnly, adLockReadOnly)

If rstSoum.EOF Then
 Call g_connData.Execute("DELETE * FROM GrbProjSoum WHERE IDProjSoum = '" & sNoCumulatif & "'")

1  Call g_connData.Execute("DELETE * FROM GrbSoumission_Pieces WHERE IDSoumission = '" & sNoCumulatif & "' AND Type = 'E'")
 
 'Efface la soumission
 Call g_connData.Execute("DELETE * FROM GrbSoumissionElec WHERE IDSoumission = '" & sNoCumulatif & "'")

 bSupprimer = True
Else
 Call rstSoumCumulatif.Open("SELECT * FROM GrbSoumissionElec WHERE IDSoumission = '" & sNoCumulatif & "'", g_connData, adOpenDynamic, adLockOptimistic)

 If rstSoum.RecordCount = 1 Then
 rstSoumCumulatif.Fields("NbreManuel") = rstSoum.Fields("NbreManuel")
 
 rstSoumCumulatif.Fields("TempsDessin") = rstSoum.Fields("TempsDessin")

 If rstSoum.Fields("SansTemps") = False Then
 rstSoumCumulatif.Fields("TempsFabrication") = rstSoum.Fields("TempsFabrication")
 Else
 rstSoumCumulatif.Fields("TempsFabrication") = 0
 End If

 rstSoumCumulatif.Fields("TempsAssemblage") = rstSoum.Fields("TempsAssemblage")
 rstSoumCumulatif.Fields("TempsProgInterface") = rstSoum.Fields("TempsProgInterface")
 rstSoumCumulatif.Fields("TempsProgAutomate") = rstSoum.Fields("TempsProgAutomate")
 rstSoumCumulatif.Fields("TempsProgRobot") = rstSoum.Fields("TempsProgRobot")
 rstSoumCumulatif.Fields("TempsVision") = rstSoum.Fields("TempsVision")
 rstSoumCumulatif.Fields("TempsTest") = rstSoum.Fields("TempsTest")
 rstSoumCumulatif.Fields("TempsInstallation") = rstSoum.Fields("TempsInstallation")
 rstSoumCumulatif.Fields("TempsMiseService") = rstSoum.Fields("TempsMiseService")
 rstSoumCumulatif.Fields("TempsFormation") = rstSoum.Fields("TempsFormation")
rstSoumCumulatif.Fields("TempsGestion") = rstSoum.Fields("TempsGestion")
 rstSoumCumulatif.Fields("TempsShipping") = rstSoum.Fields("TempsShipping")
 
 rstSoumCumulatif.Fields("NbrePersonne") = rstSoum.Fields("NbrePersonne")
 rstSoumCumulatif.Fields("TempsHebergement") = rstSoum.Fields("TempsHebergement")
 rstSoumCumulatif.Fields("TempsRepas") = rstSoum.Fields("TempsRepas")
 rstSoumCumulatif.Fields("TempsTransport") = rstSoum.Fields("TempsTransport")
 rstSoumCumulatif.Fields("TempsUniteMobile") = rstSoum.Fields("TempsUniteMobile")
 
 rstSoumCumulatif.Fields("TotalHebergement") = rstSoum.Fields("TotalHebergement")
 rstSoumCumulatif.Fields("TotalRepas") = rstSoum.Fields("TotalRepas")
 rstSoumCumulatif.Fields("PrixEmballage") = rstSoum.Fields("PrixEmballage")

 rstSoumCumulatif.Fields("total_manuel") = rstSoum.Fields("total_manuel")

 rstSoumCumulatif.Fields("MontantForfait") = rstSoum.Fields("MontantForfait")
 Else
 Do While Not rstSoum.EOF
 If Not IsNull(rstSoum.Fields("NbreManuel")) Then
 dblNbreManuel = dblNbreManuel + CDbl(rstSoum.Fields("NbreManuel"))
 End If

 If Not IsNull(rstSoum.Fields("TempsDessin")) Then
 dblTempsDessin = dblTempsDessin + CDbl(rstSoum.Fields("TempsDessin"))
 End If
 
4 If rstSoum.Fields("SansTemps") = False Then
4 If Not IsNull(rstSoum.Fields("TempsFabrication")) Then
4 dblTempsFabrication = dblTempsFabrication + CDbl(rstSoum.Fields("TempsFabrication"))
4 End If
4 End If
 
4 If Not IsNull(rstSoum.Fields("TempsAssemblage")) Then
4 dblTempsAssemblage = dblTempsAssemblage + CDbl(rstSoum.Fields("TempsAssemblage"))
4 End If
 
4 If Not IsNull(rstSoum.Fields("TempsProgInterface")) Then
4 dblTempsProgInterface = dblTempsProgInterface + CDbl(rstSoum.Fields("TempsProgInterface"))
4 End If
 
4  If Not IsNull(rstSoum.Fields("TempsProgAutomate")) Then
4  dblTempsProgAutomate = dblTempsProgAutomate + CDbl(rstSoum.Fields("TempsProgAutomate"))
4  End If
 
4  If Not IsNull(rstSoum.Fields("TempsProgRobot")) Then
4  dblTempsProgRobot = dblTempsProgRobot + CDbl(rstSoum.Fields("TempsProgRobot"))
4  End If
 
4  If Not IsNull(rstSoum.Fields("TempsVision")) Then
4  dblTempsVision = dblTempsVision + CDbl(rstSoum.Fields("TempsVision"))
50 End If
 
 If Not IsNull(rstSoum.Fields("TempsTest")) Then
 dblTempsTest = dblTempsTest + CDbl(rstSoum.Fields("TempsTest"))
 End If
 
 If Not IsNull(rstSoum.Fields("TempsInstallation")) Then
 dblTempsInstallation = dblTempsInstallation + CDbl(rstSoum.Fields("TempsInstallation"))
 End If
 
 If Not IsNull(rstSoum.Fields("TempsMiseService")) Then
 dblTempsMiseService = dblTempsMiseService + CDbl(rstSoum.Fields("TempsMiseService"))
 End If
 
 If Not IsNull(rstSoum.Fields("TempsFormation")) Then
 dblTempsFormation = dblTempsFormation + CDbl(rstSoum.Fields("TempsFormation"))
5  End If
 
5  If Not IsNull(rstSoum.Fields("TempsGestion")) Then
5  dblTempsGestion = dblTempsGestion + CDbl(rstSoum.Fields("TempsGestion"))
5  End If
 
5  If Not IsNull(rstSoum.Fields("TempsShipping")) Then
5  dblTempsShipping = dblTempsShipping + CDbl(rstSoum.Fields("TempsShipping"))
5  End If
 
5  If Not IsNull(rstSoum.Fields("TempsTransport")) Then
60 dblTempsTransport = dblTempsTransport + CDbl(rstSoum.Fields("TempsTransport"))
  End If
 
  If Not IsNull(rstSoum.Fields("TempsUniteMobile")) Then
  dblTempsUniteMobile = dblTempsUniteMobile + CDbl(rstSoum.Fields("TempsUniteMobile"))
  End If
 
  If Not IsNull(rstSoum.Fields("TotalHebergement")) Then
  dblTotalHebergement = dblTotalHebergement + CDbl(rstSoum.Fields("TotalHebergement"))
  End If
 
  If Not IsNull(rstSoum.Fields("TotalRepas")) Then
  dblTotalRepas = dblTotalRepas + CDbl(rstSoum.Fields("TotalRepas"))
  End If
 
  If Not IsNull(rstSoum.Fields("PrixEmballage")) Then
6  dblPrixEmballage = dblPrixEmballage + CDbl(rstSoum.Fields("PrixEmballage"))
6  End If
 
6  If Not IsNull(rstSoum.Fields("total_manuel")) Then
6  dblTotalManuel = dblTotalManuel + CDbl(rstSoum.Fields("total_manuel"))
6  End If

6  If Not IsNull(rstSoum.Fields("MontantForfait")) Then
6  If IsNumeric(rstSoum.Fields("MontantForfait")) Then
6  dblForfait = dblForfait + CDbl(rstSoum.Fields("MontantForfait"))
70 End If
  End If
 
  Call rstSoum.MoveNext
  Loop
 
  rstSoumCumulatif.Fields("NbreManuel") = dblNbreManuel
 
  rstSoumCumulatif.Fields("TempsDessin") = dblTempsDessin
  rstSoumCumulatif.Fields("TempsFabrication") = dblTempsFabrication
  rstSoumCumulatif.Fields("TempsAssemblage") = dblTempsAssemblage
  rstSoumCumulatif.Fields("TempsProgInterface") = dblTempsProgInterface
  rstSoumCumulatif.Fields("TempsProgAutomate") = dblTempsProgAutomate
  rstSoumCumulatif.Fields("TempsProgRobot") = dblTempsProgRobot
  rstSoumCumulatif.Fields("TempsVision") = dblTempsVision
   rstSoumCumulatif.Fields("TempsTest") = dblTempsTest
   rstSoumCumulatif.Fields("TempsInstallation") = dblTempsInstallation
7  rstSoumCumulatif.Fields("TempsMiseService") = dblTempsMiseService
7  rstSoumCumulatif.Fields("TempsFormation") = dblTempsFormation
7  rstSoumCumulatif.Fields("TempsGestion") = dblTempsGestion
7  rstSoumCumulatif.Fields("TempsShipping") = dblTempsShipping
 
7  rstSoumCumulatif.Fields("TempsTransport") = dblTempsTransport
7  rstSoumCumulatif.Fields("TempsUniteMobile") = dblTempsUniteMobile
 
80 rstSoumCumulatif.Fields("TotalHebergement") = dblTotalHebergement
  rstSoumCumulatif.Fields("TotalRepas") = dblTotalRepas
  rstSoumCumulatif.Fields("PrixEmballage") = dblPrixEmballage
 
  rstSoumCumulatif.Fields("total_manuel") = dblTotalManuel

  rstSoumCumulatif.Fields("MontantForfait") = dblForfait
  End If

  Call rstSoum.Close

  Call rstSoumCumulatif.Update

  Call rstSoumCumulatif.Close
  End If

 'AJOUT DES PIÈCES
  Call rstPiecesCumulatif.Open("SELECT * FROM GrbSoumission_Pieces WHERE IDSoumission = '" & sNoCumulatif & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
  Call g_connData.Execute("DELETE * FROM GrbSoumission_Pieces WHERE IDSoumission = '" & sNoCumulatif & "'")

   Call rstPieces.Open("SELECT * FROM GrbSoumission_Pieces WHERE LEFT(IDSoumission, 6) = '" & Left$(txtNoProjSoum.Text, 6) & "' AND RIGHT(IDSoumission, 2) <> '99' ORDER BY CInt(OrdreSection), NuméroLigne", g_connData, adOpenForwardOnly, adLockReadOnly)

   Do While Not rstPieces.EOF
   Call rstPiecesCumulatif.AddNew

   rstPiecesCumulatif.Fields("IDSoumission") = sNoCumulatif
8  rstPiecesCumulatif.Fields("IDSection") = rstPieces.Fields("IDSection")
8  rstPiecesCumulatif.Fields("NumItem") = rstPieces.Fields("NumItem")
8  rstPiecesCumulatif.Fields("Qté") = rstPieces.Fields("Qté")
8  rstPiecesCumulatif.Fields("Desc_FR") = rstPieces.Fields("Desc_FR")
90 rstPiecesCumulatif.Fields("Desc_EN") = rstPieces.Fields("Desc_EN")
  rstPiecesCumulatif.Fields("Manufact") = rstPieces.Fields("Manufact")
  rstPiecesCumulatif.Fields("Prix_list") = rstPieces.Fields("Prix_list")
  rstPiecesCumulatif.Fields("Escompte") = rstPieces.Fields("Escompte")
  rstPiecesCumulatif.Fields("Prix_net") = rstPieces.Fields("Prix_net")
  rstPiecesCumulatif.Fields("IDFRS") = rstPieces.Fields("IDFRS")
  rstPiecesCumulatif.Fields("Temps") = rstPieces.Fields("Temps")
  rstPiecesCumulatif.Fields("Temps_Total") = rstPieces.Fields("Temps_Total")
  rstPiecesCumulatif.Fields("Prix_total") = rstPieces.Fields("Prix_total")
  rstPiecesCumulatif.Fields("Profit_Argent") = rstPieces.Fields("Profit_Argent")
  rstPiecesCumulatif.Fields("SousSection") = rstPieces.Fields("SousSection")
  rstPiecesCumulatif.Fields("OrdreSection") = rstPieces.Fields("OrdreSection")
 rstPiecesCumulatif.Fields("NuméroLigne") = rstPieces.Fields("NuméroLigne")
   rstPiecesCumulatif.Fields("PrixOrigine") = rstPieces.Fields("PrixOrigine")
 rstPiecesCumulatif.Fields("Type") = rstPieces.Fields("Type")
   rstPiecesCumulatif.Fields("Visible") = rstPieces.Fields("Visible")
 rstPiecesCumulatif.Fields("Commentaire") = rstPieces.Fields("Commentaire")
   rstPiecesCumulatif.Fields("Devise") = rstPieces.Fields("Devise")
 rstPiecesCumulatif.Fields("Provenance") = Right(rstPieces.Fields("IDSoumission"), 2)

9  Call rstPiecesCumulatif.Update

 Call rstPieces.MoveNext
100 Loop

10 Call rstPiecesCumulatif.Close
10 Call rstPieces.Close

Set rstSoum = New ADODB.Recordset
10 Set rstPieces = New ADODB.Recordset
Set rstSoumCumulatif = New ADODB.Recordset
10 Set rstPiecesCumulatif = New ADODB.Recordset

If bSupprimer = False Then
1 Call CalculerTotalRecordset(sNoCumulatif)
End If

10 Exit Sub

Oups:

10  wOups "FrmProjSoumElec", "RecreerSoumissionCumulatif", Err, Err.number, Err.Description
End Sub
Private Function ExportdansExcel(ByVal oRecordset As ADODB.Recordset)

 On Error GoTo Oups

  Dim iCount As Integer
 Dim oXLApp As Excel.Application 'Declare the object variables
 Dim oXLBook As Excel.Workbook
 Dim oXLSheet As Excel.Worksheet

 Set oXLApp = New Excel.Application 'Create a new instance of Excel
 Set oXLBook = oXLApp.Workbooks.Add 'Add a new workbook
 Set oXLSheet = oXLBook.Worksheets(1) 'Work with the first worksheet

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

oXLSheet.range("A1: N1").Font.Bold = True


 With oXLSheet 'Fill with data
 For iCount = 0 To (oRecordset.Fields.count - 1)
 .Cells(1, iCount + 1) = oRecordset.Fields(iCount).Name
 Next iCount
  'create and fill a recordset here, called oRecordset
  .range("A2").CopyFromRecordset oRecordset
  End With

  oXLApp.Visible = True 'Show it to the user
  Set oXLSheet = Nothing 'Disconnect from all Excel objects (let the user take over)
  Set oXLBook = Nothing
  Set oXLApp = Nothing

  Exit Function

Oups:

 wOups "FrmProjSoumMec", "ExportdansExcel", Err, Err.number, Err.Description
End Function

