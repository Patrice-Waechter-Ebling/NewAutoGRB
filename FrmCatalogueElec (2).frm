VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCatalogueElec 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catalogue Électrique"
   ClientHeight    =   7065
   ClientLeft      =   1455
   ClientTop       =   1020
   ClientWidth     =   10395
   Icon            =   "FrmCatalogueElec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   10395
   Begin MSComctlLib.ListView lvwCategorie 
      Height          =   2295
      Left            =   5280
      TabIndex        =   86
      Top             =   360
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Catalogue"
         Object.Width           =   4471
      EndProperty
   End
   Begin VB.CommandButton CmdRecherchecategorie 
      Height          =   375
      Left            =   9960
      Picture         =   "FrmCatalogueElec.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   360
      Width           =   375
   End
   Begin MSComctlLib.ListView lvwRechercheAchat 
      Height          =   2295
      Left            =   1560
      TabIndex        =   83
      Top             =   960
      Visible         =   0   'False
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No. Job"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nbre fois"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvwRechercheJob 
      Height          =   2175
      Left            =   1560
      TabIndex        =   21
      Top             =   1080
      Visible         =   0   'False
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No. Job"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nbre fois"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdCopier 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Copier"
      Height          =   495
      Left            =   3360
      TabIndex        =   73
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton cmdRechercheInventaire 
      Caption         =   "Inventaire"
      Height          =   495
      Left            =   120
      TabIndex        =   71
      Top             =   6480
      Width           =   1455
   End
   Begin MSComctlLib.ListView lvwFabricant 
      Height          =   2295
      Left            =   1560
      TabIndex        =   13
      Top             =   960
      Visible         =   0   'False
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
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
         Text            =   "Manufacturier"
         Object.Width           =   2090
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "No. Pièce"
         Object.Width           =   3254
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description française"
         Object.Width           =   6138
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Description anglaise"
         Object.Width           =   6138
      EndProperty
   End
   Begin MSComctlLib.ListView lvwPieces 
      Height          =   2295
      Left            =   1560
      TabIndex        =   11
      Top             =   960
      Visible         =   0   'False
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
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
         Text            =   "No Pièce"
         Object.Width           =   3254
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Manufacturier"
         Object.Width           =   2090
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description française"
         Object.Width           =   6138
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Description anglaise"
         Object.Width           =   6138
      EndProperty
   End
   Begin MSComctlLib.ListView lvwDescription 
      Height          =   2295
      Left            =   1560
      TabIndex        =   12
      Top             =   960
      Visible         =   0   'False
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
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
         Text            =   "Description française"
         Object.Width           =   6138
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description anglaise"
         Object.Width           =   6138
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Manufacturier"
         Object.Width           =   2090
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "No Pièce"
         Object.Width           =   3254
      EndProperty
   End
   Begin VB.CommandButton cmdRechercherPiece 
      Height          =   375
      Left            =   3600
      Picture         =   "FrmCatalogueElec.frx":0784
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton cmdChangerCategorie 
      Caption         =   "Changer de catégorie"
      Height          =   375
      Left            =   7920
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdDemande 
      Caption         =   "Demande de prix"
      Height          =   495
      Left            =   1680
      TabIndex        =   72
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton cmdRechercheDescrFR 
      Height          =   375
      Left            =   9960
      Picture         =   "FrmCatalogueElec.frx":0886
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   900
      Width           =   375
   End
   Begin VB.TextBox txtTemps 
      BackColor       =   &H00FFFFFF&
      DataField       =   "TEMPS"
      DataSource      =   "DatCat1"
      Height          =   285
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   2040
      Width           =   735
   End
   Begin VB.Frame frafournisseur 
      BackColor       =   &H00000000&
      Caption         =   "Fournisseur"
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   120
      TabIndex        =   45
      Top             =   3960
      Width           =   10215
      Begin VB.TextBox txtTauxChange 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   8640
         TabIndex        =   80
         Top             =   1920
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSComctlLib.ListView lvwfournisseur 
         Height          =   1575
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   9975
         _ExtentX        =   17595
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
         NumItems        =   10
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
            SubItemIndex    =   5
            Text            =   "Prix listé"
            Object.Width           =   1614
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Escompte"
            Object.Width           =   1561
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Prix net"
            Object.Width           =   1614
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Prix spécial"
            Object.Width           =   1720
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Quoter"
            Object.Width           =   1191
         EndProperty
      End
      Begin VB.CommandButton cmdAddFrs 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Ajouter"
         Height          =   375
         Left            =   120
         TabIndex        =   66
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdSuppFrs 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Supprimer"
         Height          =   375
         Left            =   1320
         TabIndex        =   69
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdModifFrs 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Modifier"
         Height          =   375
         Left            =   2520
         TabIndex        =   70
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CheckBox chkquoter 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Quoter :"
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
         TabIndex        =   53
         Top             =   1440
         Width           =   1335
      End
      Begin VB.OptionButton optUSA 
         BackColor       =   &H00000000&
         Caption         =   "USA"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   8640
         TabIndex        =   63
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton optCAN 
         BackColor       =   &H00000000&
         Caption         =   "CAN"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7920
         TabIndex        =   62
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.TextBox txtPrixList 
         BackColor       =   &H00FFFFFF&
         DataField       =   "PRIX_LIST"
         DataSource      =   "DatCat1"
         Height          =   285
         Left            =   6360
         TabIndex        =   55
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtPrixNet 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6360
         TabIndex        =   59
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtPrixSpecial 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6360
         TabIndex        =   61
         Top             =   1440
         Width           =   855
      End
      Begin MSMask.MaskEdBox mskValide 
         Height          =   255
         Left            =   1560
         TabIndex        =   52
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskEscompte 
         Height          =   255
         Left            =   6360
         TabIndex        =   57
         Top             =   720
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.CommandButton cmdEnrFrs 
         Caption         =   "&Enregistre"
         Height          =   375
         Left            =   120
         TabIndex        =   67
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdAnnulFrs 
         Caption         =   "A&nnuler"
         Height          =   375
         Left            =   1320
         TabIndex        =   68
         Top             =   1920
         Width           =   1095
      End
      Begin VB.ComboBox cmbfrs 
         Height          =   315
         Left            =   1560
         TabIndex        =   48
         Top             =   360
         Width           =   2775
      End
      Begin VB.ComboBox cmbPersRess 
         Height          =   315
         Left            =   1560
         TabIndex        =   50
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton optSpain 
         BackColor       =   &H00000000&
         Caption         =   "SPA"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   9360
         TabIndex        =   64
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblDevise2 
         BackStyle       =   0  'Transparent
         Caption         =   "$ USA"
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
         Left            =   9550
         TabIndex        =   82
         Top             =   1920
         Visible         =   0   'False
         Width           =   575
      End
      Begin VB.Label lblDevise1 
         BackStyle       =   0  'Transparent
         Caption         =   "1$ CAN ="
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
         Left            =   7800
         TabIndex        =   81
         Top             =   1920
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   2880
         TabIndex        =   65
         Top             =   1080
         Width           =   975
      End
      Begin VB.Image imgSpain 
         Height          =   1065
         Left            =   8160
         Picture         =   "FrmCatalogueElec.frx":0BC8
         Stretch         =   -1  'True
         Top             =   720
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Image imgCanada 
         Height          =   1065
         Left            =   8160
         Picture         =   "FrmCatalogueElec.frx":64BB2
         Stretch         =   -1  'True
         Top             =   720
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Image imgEU 
         Height          =   1065
         Left            =   8160
         Picture         =   "FrmCatalogueElec.frx":BAB94
         Stretch         =   -1  'True
         Top             =   720
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
         Left            =   120
         TabIndex        =   47
         Top             =   360
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
         Left            =   5160
         TabIndex        =   54
         Top             =   360
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
         Left            =   5160
         TabIndex        =   56
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pers. Ress :"
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
         Index           =   24
         Left            =   120
         TabIndex        =   49
         Top             =   720
         Width           =   1575
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
         Left            =   5160
         TabIndex        =   58
         Top             =   1080
         Width           =   1095
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
         Index           =   2
         Left            =   5160
         TabIndex        =   60
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Valide jusqu'au :"
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
         Index           =   23
         Left            =   120
         TabIndex        =   51
         Top             =   1080
         Width           =   1455
      End
   End
   Begin VB.ComboBox cmbCategorie 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5280
      TabIndex        =   3
      Text            =   "cmbCategorie"
      Top             =   360
      Width           =   2535
   End
   Begin VB.TextBox txtNoItemGRB 
      BackColor       =   &H00FFFFFF&
      DataField       =   "PIECE_GRB"
      DataSource      =   "DatCat1"
      Height          =   285
      Left            =   1560
      MaxLength       =   21
      TabIndex        =   22
      Top             =   2280
      Width           =   1935
   End
   Begin VB.ComboBox cmbNoItem 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1560
      TabIndex        =   16
      Text            =   "cmbNoItem"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton CmdModif 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Modifier"
      Height          =   495
      Left            =   7680
      TabIndex        =   78
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton CmdFerme 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Fermer"
      Height          =   495
      Left            =   9120
      TabIndex        =   79
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton CmdSupp 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Supprimer"
      Height          =   495
      Left            =   6240
      TabIndex        =   76
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton CmdAdd 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Ajouter"
      Height          =   495
      Left            =   4800
      TabIndex        =   74
      Top             =   6480
      Width           =   1335
   End
   Begin VB.TextBox txtPageCat 
      BackColor       =   &H00FFFFFF&
      DataField       =   "PAGE_CAT"
      DataSource      =   "DatCat1"
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   28
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox txtComment 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5640
      Locked          =   -1  'True
      MaxLength       =   41
      TabIndex        =   25
      Top             =   2520
      Width           =   2415
   End
   Begin VB.ComboBox cmbFabricant 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Text            =   "cmbFabricant"
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox txtFabricant 
      BackColor       =   &H80000016&
      DataField       =   "Manufact"
      DataSource      =   "DatCat1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtDescriptionEN 
      BackColor       =   &H00FFFFFF&
      DataField       =   "DESCR_EN"
      DataSource      =   "DatCat1"
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      MaxLength       =   61
      TabIndex        =   20
      Top             =   1440
      Width           =   4575
   End
   Begin VB.TextBox txtDescriptionFR 
      BackColor       =   &H00FFFFFF&
      DataField       =   "DESCR_FR"
      DataSource      =   "DatCat1"
      Height          =   315
      Left            =   5280
      Locked          =   -1  'True
      MaxLength       =   61
      TabIndex        =   9
      Top             =   960
      Width           =   4575
   End
   Begin VB.TextBox txtNoItem 
      DataField       =   "PIECE"
      DataSource      =   "DatCat1"
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1800
      Width           =   1932
   End
   Begin VB.CommandButton CmdAnul 
      Caption         =   "A&nnuler"
      Height          =   495
      Left            =   6240
      TabIndex        =   77
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton CmdEnr 
      Caption         =   "&Enregistrer"
      Default         =   -1  'True
      Height          =   495
      Left            =   4800
      TabIndex        =   75
      Top             =   6480
      Width           =   1335
   End
   Begin VB.TextBox txtCategorie 
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton cmdRechercherManufacturier 
      Height          =   375
      Left            =   3600
      Picture         =   "FrmCatalogueElec.frx":107906
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   960
      Width           =   375
   End
   Begin VB.CheckBox chkInventaire 
      BackColor       =   &H00000000&
      Caption         =   "Présent dans l'inventaire"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8160
      TabIndex        =   26
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Frame fraQuantité 
      BackColor       =   &H00000000&
      Caption         =   "Quantité"
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
      Height          =   1695
      Left            =   8160
      TabIndex        =   36
      Top             =   2160
      Width           =   2175
      Begin VB.CheckBox chkBoite 
         BackColor       =   &H00000000&
         Caption         =   "Par Boîte :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtQuantitéBoite 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   38
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtQuantiteMinimum 
         Height          =   285
         Left            =   1320
         TabIndex        =   42
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtQuantiteStock 
         Height          =   285
         Left            =   1320
         TabIndex        =   40
         Top             =   600
         Width           =   615
      End
      Begin VB.CheckBox chkMinimum 
         BackColor       =   &H00000000&
         Caption         =   "Minimum :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtQuantiteCommande 
         Height          =   285
         Left            =   1320
         TabIndex        =   44
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "À commander :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   1320
         Width           =   1095
      End
   End
   Begin VB.ComboBox cmbLocalisation 
      Height          =   315
      Left            =   6480
      TabIndex        =   30
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox txtLocalisation 
      Height          =   285
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   3480
      Width           =   1575
   End
   Begin VB.ComboBox cmbDescriptionFR 
      Height          =   315
      Left            =   5280
      TabIndex        =   10
      Top             =   960
      Width           =   4575
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      Height          =   375
      Left            =   120
      TabIndex        =   32
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdRechercheJob 
      Caption         =   "Recherche dans jobs / soums"
      Height          =   495
      Left            =   1080
      TabIndex        =   33
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdRechercheAchat 
      Caption         =   "Recherche dans achats"
      Height          =   495
      Left            =   2520
      TabIndex        =   84
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Localisation :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6480
      TabIndex        =   29
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Temps :"
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
      Left            =   4200
      TabIndex        =   34
      Top             =   2040
      Width           =   1305
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Catégorie de pièce :"
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
      Left            =   3360
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pièce GRB :"
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
      Index           =   25
      Left            =   120
      TabIndex        =   23
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Catalogue :"
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
      TabIndex        =   27
      Top             =   2880
      Width           =   1335
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
      Index           =   6
      Left            =   4200
      TabIndex        =   24
      Top             =   2520
      Width           =   1305
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Desc. EN :"
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
      Left            =   4080
      TabIndex        =   19
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Desc. FR :"
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
      Left            =   4080
      TabIndex        =   8
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacturier :"
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
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pièce :"
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
      TabIndex        =   15
      Top             =   1800
      Width           =   1335
   End
End
Attribute VB_Name = "FrmCatalogueElec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  
'Numéros de colonne du ListView pour la recherche par description
Private Const I_COL_DES_DESCR_FR As Integer = 0
Private Const I_COL_DES_DESCR_EN As Integer = 1
Private Const I_COL_DES_FABRICANT As Integer = 2
Private Const I_COL_DES_PIECE As Integer = 3

'Numéros de colonne du ListView pour la recherche par pièce
Private Const I_COL_PIECE_PIECE As Integer = 0
Private Const I_COL_PIECE_FABRICANT As Integer = 1
Private Const I_COL_PIECE_DESCR_FR As Integer = 2
Private Const I_COL_PIECE_DESCR_EN As Integer = 3

'Numéros de colonne du ListView pour la recherche par manufacturier
Private Const I_COL_MAN_FABRICANT As Integer = 0
Private Const I_COL_MAN_PIECE As Integer = 1
Private Const I_COL_MAN_DESCR_FR As Integer = 2
Private Const I_COL_MAN_DESCR_EN As Integer = 3

'Numéros de colonne du ListView pour les fournisseurs
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

'Numéro de colonne du ListView pour la recherche dans les jobs
Private Const I_COL_JOB_NUMERO As Integer = 0
Private Const I_COL_JOB_QUANTITE As Integer = 1

Private Const I_COL_ACHAT_NUMERO As Integer = 0
Private Const I_COL_ACHAT_QUANTITE As Integer = 1

Public Enum enumModeCatalogueElec
 MODE_AJOUT_MODIF_ELEC = 0
 MODE_INACTIF = 1
 MODE_AJOUT_MODIF_FRS = 2
End Enum

Public m_eDemande As enumModeDemande
Public m_bDemandeAnnuler As Boolean
Public m_bAjout As Boolean
Public m_bAnnulerCopie As Boolean
Public m_sCategorieCopie As String
Public m_bPieceEffacée As Boolean
Private m_bRempliManuel As Boolean
Private m_sNoItem As String
Private m_eMode As enumModeCatalogueElec
Private m_bBloqueDescription As Boolean
Private m_collPieceDescFR As Collection

'Pour pouvoir comparer la quantité stock avant et après une modification
'pour savoir que c'est de l'ajustement d'inventaire
Private m_sQteStockAvant As String

'Pour pouvoir choisir lors du remplissage
Public m_sSelectCategorie As String
Public m_sSelectFabricant As String
Public m_sSelectNoItem As String

Private m_bCopiePiece As Boolean
'utilisé pour créer la condition pour les recordsets si on choisi tous les fabricant
Public sChoisirTous As String


Public Sub ViderChamps_frs()

 On Error GoTo Oups

 'Enlever la sélection dans le combo
 cmbfrs.ListIndex = -1

 'Vide les champs pieces
 txtPrixSpecial.Text = vbNullString
 cmbPersRess.ListIndex = -1
 txtPrixList.Text = vbNullString
 mskEscompte.Text = vbNullString
 txtPrixNet.Text = vbNullString
 mskValide.Text = vbNullString
 
 'Enlève le check
 chkquoter.Value = vbUnchecked
 optCAN.Value = True

 Exit Sub

Oups:

  wOups "frmCatalogueElec", "ViderChamps_frs", Err, Err.number, Err.Description
End Sub

Public Sub ViderChamps_piece()

 On Error GoTo Oups

 'Vide les champs pieces
 txtNoItemGRB.Text = vbNullString
 txtDescriptionEN.Text = vbNullString
 txtTemps.Text = vbNullString
 txtComment.Text = vbNullString
 txtQuantitéBoite.Text = vbNullString
 txtQuantiteCommande.Text = vbNullString
 txtQuantiteMinimum.Text = vbNullString
 txtQuantiteStock.Text = vbNullString
 txtLocalisation.Text = vbNullString

 cmbLocalisation.ListIndex = -1
 
 'Enlève le check
  chkBoite.Value = vbUnchecked
  chkInventaire.Value = vbUnchecked
  chkMinimum.Value = vbUnchecked

  Exit Sub

Oups:

  wOups "frmCatalogueElec", "ViderChamps_piece", Err, Err.number, Err.Description
End Sub

Public Sub BarrerChamps_piece(ByVal bLocked As Boolean)

 On Error GoTo Oups

 'Barre les champs
 txtNoItem.Locked = bLocked
 txtNoItemGRB.Locked = bLocked
 txtDescriptionEN.Locked = bLocked
 txtDescriptionFR.Locked = bLocked
 txtTemps.Locked = bLocked
  txtComment.Locked = bLocked
  frafournisseur.Enabled = bLocked
  chkInventaire.Enabled = Not bLocked

  If chkInventaire.Enabled = True Then
  If chkInventaire.Value = vbChecked Then
  fraQuantité.Enabled = True
  Else
  fraQuantité.Enabled = False
End If
Else
 fraQuantité.Enabled = False
End If

Exit Sub

Oups:

wOups "frmCatalogueElec", "BarrerChamps_piece", Err, Err.number, Err.Description
End Sub

Public Sub MontrerControles(ByVal eMode As enumModeCatalogueElec)

 On Error GoTo Oups

 'Mets des champs visible et d'autres invisible
 Dim bCategorie As Boolean
 Dim bFabricant As Boolean
 Dim bNoItem As Boolean
 Dim bLocalisation As Boolean
 Dim bCmdAddFRS As Boolean
 Dim bCmdModifFRS As Boolean
 Dim bCmdSuppFRS As Boolean
 Dim bCmdEnrFRS As Boolean
 Dim bCmdAnnulFRS As Boolean
 Dim bCmdAdd As Boolean
  Dim bCmdModif As Boolean
  Dim bCmdSupp As Boolean
  Dim bCmdFermer As Boolean
  Dim bCmdEnr As Boolean
  Dim bCmdAnnul As Boolean
  Dim bFraFRS As Boolean
  Dim bLvwFRS As Boolean
  Dim bCmdSearchMan As Boolean
10 Dim bCmdSearchPiece As Boolean
Dim bCmdSearchDescr As Boolean
Dim bCmdDemande As Boolean
Dim bCmbDescFR As Boolean
Dim bCopier As Boolean
Dim bChangerCat As Boolean
Dim bInventaire As Boolean

m_eMode = eMode
 
Select Case eMode
 Case MODE_INACTIF:
 bCategorie = True
 bFabricant = True
 bNoItem = True
 bCmdAddFRS = True
 bCmdModifFRS = True
 bCmdSuppFRS = True
 bCmdAdd = True
 bCmdModif = True
 bCmdSupp = True
 bCmdFermer = True
1  bFraFRS = True
 bLvwFRS = True
 bCmdSearchMan = True
 bCmdSearchPiece = True
 bCmdSearchDescr = True
 bCmdDemande = True
 bCopier = True
 bInventaire = True
 bCmbDescFR = True
 
 Case MODE_AJOUT_MODIF_ELEC:
 bCmdAddFRS = True
 bCmdModifFRS = True
 bCmdSuppFRS = True
 bCmdEnr = True
 bFabricant = True 'GLL 2017-09-01
 txtFabricant.Enabled = True
 bCmdAnnul = True
 bLvwFRS = True
 bCmdSearchDescr = True
 bLocalisation = True
 bChangerCat = True
 
Case MODE_AJOUT_MODIF_FRS:
 bCmdEnrFRS = True
 bCmdAnnulFRS = True
bFraFRS = True
End Select
 
cmbCategorie.Visible = bCategorie
txtCategorie.Visible = Not bCategorie
 
cmbDescriptionFR.Visible = bCmbDescFR
txtDescriptionFR.Visible = Not bCmbDescFR
 
cmbFabricant.Visible = bFabricant
txtFabricant.Visible = bFabricant
 
cmbNoItem.Visible = bNoItem
txtNoItem.Visible = Not bNoItem
 
cmbLocalisation.Visible = bLocalisation
3  txtLocalisation.Visible = Not bLocalisation
 
frafournisseur.Enabled = bFraFRS
 
3  lvwfournisseur.Visible = bLvwFRS
 
cmdAddFrs.Visible = bCmdAddFRS
3  cmdModifFrs.Visible = bCmdModifFRS
cmdSuppFrs.Visible = bCmdSuppFRS
3  cmdEnrFrs.Visible = bCmdEnrFRS
 cmdAnnulFrs.Visible = bCmdAnnulFRS
40 CmdAdd.Visible = bCmdAdd
CmdModif.Visible = bCmdModif
4 CmdSupp.Visible = bCmdSupp
4 CmdFerme.Visible = bCmdFermer
4 CmdEnr.Visible = bCmdEnr
4 CmdAnul.Visible = bCmdAnnul
4 cmdDemande.Visible = bCmdDemande
4 cmdCopier.Visible = bCopier
4 cmdRechercheDescrFR.Enabled = bCmdSearchDescr
4 cmdRechercherPiece.Enabled = bCmdSearchPiece
4 cmdRechercherManufacturier.Enabled = bCmdSearchMan
4 cmdChangerCategorie.Visible = bChangerCat
4  cmdRechercheInventaire.Visible = bInventaire


4  lblDevise1.Visible = False
4  txtTauxChange.Visible = False
4  lblDevise2.Visible = False

4  Exit Sub

Oups:

4  wOups "frmCatalogueElec", "MontrerControles", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboPersRess()

 On Error GoTo Oups

 Dim rstContactFRS As ADODB.Recordset
 Dim rstContact As ADODB.Recordset
 
 Call cmbPersRess.Clear

 Set rstContactFRS = New ADODB.Recordset
 
 Call rstContactFRS.Open("SELECT * FROM GrbContactFRS WHERE NoFRS = " & cmbfrs.ItemData(cmbfrs.ListIndex), g_connData, adOpenDynamic, adLockOptimistic)
 
 Set rstContact = New ADODB.Recordset
 
 Do While Not rstContactFRS.EOF
 Call rstContact.Open("SELECT IDContact, NomContact FROM GrbContact WHERE IDContact = " & rstContactFRS.Fields("NoContact"), g_connData, adOpenDynamic, adLockOptimistic)
 
 If Not rstContact.EOF Then
 Call cmbPersRess.AddItem(rstContact.Fields("NomContact"))
 
  cmbPersRess.ItemData(cmbPersRess.newIndex) = rstContact.Fields("IDContact")
  End If
 
  Call rstContact.Close

  Call rstContactFRS.MoveNext
  Loop

  If cmbPersRess.ListCount = 0 Then
  Call rstContact.Open("SELECT NomContact, IDContact FROM GrbContact WHERE Supprimé = False ORDER BY NomContact", g_connData, adOpenDynamic, adLockOptimistic)
 
  Do While Not rstContact.EOF
 Call cmbPersRess.AddItem(rstContact.Fields("NomContact"))
 
cmbPersRess.ItemData(cmbPersRess.newIndex) = rstContact.Fields("IDContact")
 
 Call rstContact.MoveNext
 Loop
 
 Call rstContact.Close
End If

Set rstContact = Nothing

Exit Sub

Oups:

wOups "frmCatalogueElec", "RemplirComboPersRess", Err, Err.number, Err.Description
End Sub

Private Sub chkBoite_Click()
 
 On Error GoTo Oups

 If m_eMode = MODE_AJOUT_MODIF_ELEC Then
 If chkBoite.Value = vbChecked Then
 txtQuantitéBoite.Enabled = True
 Else
 txtQuantitéBoite.Enabled = False
 End If
 End If

 Exit Sub

Oups:
 
 wOups "frmCatalogueElec", "chkBoite_Click", Err, Err.number, Err.Description
End Sub

Private Sub chkInventaire_Click()

 On Error GoTo Oups

 If m_eMode = MODE_AJOUT_MODIF_ELEC Then
 If chkInventaire.Value = vbChecked Then
 fraQuantité.Enabled = True
 cmbLocalisation.Enabled = True
 Else
 fraQuantité.Enabled = False
 cmbLocalisation.Enabled = False
 End If
 End If

 Exit Sub

Oups:

  wOups "frmCatalogueElec", "chkInventaire_Click", Err, Err.number, Err.Description
End Sub

Private Sub chkMinimum_Click()

 On Error GoTo Oups

 If m_eMode = MODE_AJOUT_MODIF_ELEC Then
 If chkMinimum.Value = vbChecked Then
 txtQuantiteMinimum.Enabled = True
 txtQuantiteCommande.Enabled = True
 Else
 txtQuantiteMinimum.Enabled = False
 txtQuantiteCommande.Enabled = False
 End If
 End If

 Exit Sub

Oups:

  wOups "frmCatalogueElec", "chkMinimum_Click", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboDescription()

 On Error GoTo Oups

 'Remplir le combo des descriptions
 Dim rstCatElec As ADODB.Recordset
 Dim sPiece As String
 Dim sCategorie As String
Dim sFabricant As String

 Do While m_collPieceDescFR.count > 0
 Call m_collPieceDescFR.Remove(1)
 Loop
 
 Call cmbDescriptionFR.Clear

 sCategorie = Replace(cmbCategorie.Text, "'", "''")
4  sFabricant = Replace(cmbFabricant.Text, "'", "''")

 Set rstCatElec = New ADODB.Recordset
 
4 If sFabricant = "-- CHOISIR TOUS --" Then
 If cmbCategorie.Text = "DIVERS" Or sChoisirTous = ")" Then
 Call rstCatElec.Open("SELECT * FROM GrbCatalogueElec WHERE CATEGORIE = '" & sCategorie & "' ORDER BY TRIM(PIECE)", g_connData, adOpenDynamic, adLockOptimistic)
54 Else
 Call rstCatElec.Open("SELECT * FROM GrbCatalogueElec WHERE CATEGORIE = '" & sCategorie & "'" & sChoisirTous & " ORDER BY TRIM(PIECE)", g_connData, adOpenDynamic, adLockOptimistic)
 End If
 Else
 
5  Call rstCatElec.Open("SELECT * FROM GrbCatalogueElec WHERE CATEGORIE = '" & sCategorie & "' AND FABRICANT = '" & sFabricant & "' ORDER BY TRIM(PIECE)", g_connData, adOpenDynamic, adLockOptimistic)
57
 End If

 
  Do While Not rstCatElec.EOF
  If Not IsNull(rstCatElec.Fields("DESC_FR")) Then
  If rstCatElec.Fields("DESC_FR") <> vbNullString Then
  Call cmbDescriptionFR.AddItem(Trim(rstCatElec.Fields("DESC_FR")))
 
  sPiece = Trim(rstCatElec.Fields("PIECE"))
 
  Call m_collPieceDescFR.Add(sPiece)
  End If
  End If
 
Call rstCatElec.MoveNext
Loop
 
Call rstCatElec.Close
Set rstCatElec = Nothing

Exit Sub

Oups:

wOups "frmCatalogueElec", "RemplirComboDescription", Err, Err.number, Err.Description
End Sub

Private Sub cmbDescriptionFR_Click()

 On Error GoTo Oups

 Dim rstCatElec As ADODB.Recordset
 Dim sNoItem As String
 Dim sFabricant As String
 Dim iCompteur As Integer

 txtDescriptionFR.Text = cmbDescriptionFR.Text

 If m_bBloqueDescription = False Then
 For iCompteur = 0 To cmbNoItem.ListCount - 1
 If cmbNoItem.LIST(iCompteur) = m_collPieceDescFR(cmbDescriptionFR.ListIndex + 1) Then
 cmbNoItem.ListIndex = iCompteur

 Exit For
  End If
  Next
  End If

  Exit Sub

Oups:

  wOups "frmCatalogueElec", "cmbDescriptionFR_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmbfrs_Click()

 On Error GoTo Oups

 If cmbfrs.ListIndex <> -1 Then
 cmbfrs.Tag = cmbfrs.ItemData(cmbfrs.ListIndex)
 
 Call RemplirComboPersRess
 End If

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "cmbfrs_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmbLocalisation_Click()

 On Error GoTo Oups

 txtLocalisation.Text = cmbLocalisation.Text

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "cmbLocalisation_Click", Err, Err.number, Err.Description
End Sub

Private Sub CmdAdd_Click()

 On Error GoTo Oups

 'Montre le dialogue pour ajouter un item au catalogue
 Screen.MousePointer = vbHourglass
 
 m_bBloqueDescription = True
 
 Call OuvrirForm(FrmaddItemElec, True)

 m_bBloqueDescription = False
 
 Screen.MousePointer = vbDefault

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "CmdAdd_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdAddFrs_Click()

 On Error GoTo Oups

 If cmbNoItem.ListCount > 0 Then
 'ajoute un fournisseur pour la piece
 m_bAjout = True

 Call BarrerChamps_piece(True)

 Call ViderChamps_frs

 Call cmbfrs.SetFocus

 Call MontrerControles(MODE_AJOUT_MODIF_FRS)
 
 'affiche drapeau
 Call AfficherDrapeau
 End If

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "cmdAddFrs_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnulFrs_Click()

 On Error GoTo Oups

 Call MontrerControles(MODE_INACTIF)

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "cmdAnnulFrs_Click", Err, Err.number, Err.Description
End Sub

Private Sub CmdAnul_Click()

 On Error GoTo Oups

 txtPrixNet.Enabled = True
 txtPrixSpecial.Enabled = True

 m_bBloqueDescription = True
 txtFabricant.Top = 1320 'GLL 2017-10-10 Modification pour ajouter la modification possible du Manufacturier
 cmbFabricant.Visible = True 'GLL 2017-10-10 Modification pour ajouter la modification possible du Manufacturier
 Call AfficherItem
 
 m_bBloqueDescription = False

 m_bCopiePiece = False
 
 'on cache les combos
 cmbFabricant.Visible = False
 cmbNoItem.Visible = False

 'on retablis les boutons
 Call MontrerControles(MODE_INACTIF)
 Call BarrerChamps_piece(True)

  m_sQteStockAvant = ""

  Exit Sub

Oups:

  wOups "frmCatalogueElec", "CmdAnul_Click", Err, Err.number, Err.Description
End Sub

Private Sub EnregistrerItem()

 On Error GoTo Oups

 'Enregistrement de l'item dans la BD
 Dim rstItem As ADODB.Recordset
 Dim rstItemFRS As ADODB.Recordset
 Dim rstItemFRSDest As ADODB.Recordset
 Dim rstVerif As ADODB.Recordset
 Dim rstInventaire As ADODB.Recordset
 Dim rstInvModif As ADODB.Recordset
 Dim sNomFab As String
 Dim sNoPiece As String
 Dim iCompteur As Integer
 Dim sPieceModif As String
  Dim sLettre As String
 
  sNomFab = txtFabricant.Text
  sNoPiece = txtNoItem.Text
 
  If m_bCopiePiece = True Or (m_bCopiePiece = False And (UCase(sNoPiece) <> UCase(m_sNoItem))) Then
  Set rstVerif = New ADODB.Recordset

  Call rstVerif.Open("SELECT * FROM GrbCatalogueElec WHERE PIECE = '" & Replace(sNoPiece, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)

  If Not rstVerif.EOF Then
  Call MsgBox("Le numéro " & sNoPiece & " existe déjà!", vbOKOnly, "Erreur")

 Call rstVerif.Close
Set rstVerif = Nothing

 Exit Sub
 End If

 Call rstVerif.Close
 Set rstVerif = Nothing
End If
 
If txtFabricant.Text = vbNullString Or txtNoItem.Text = vbNullString Or txtDescriptionFR.Text = vbNullString Then
 Call MsgBox("Les champs Manufacturier, Pièce et Desc. FR doivent être remplis!", vbOKOnly, "Erreur")
 
 Exit Sub
End If
 
 'Sinon, j'ouvre un recordset contenant le no d'item
Set rstItem = New ADODB.Recordset
 
1  Call rstItem.Open("SELECT * FROM GrbCatalogueElec WHERE PIECE = '" & Replace(m_sNoItem, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 'enregistre le nopiece dans la table distributeur si pas vide
Set rstItemFRS = New ADODB.Recordset
 
 Call rstItemFRS.Open("SELECT * FROM GrbPiecesFRS WHERE PIECE = '" & Replace(rstItem.Fields("PIECE"), "'", "''") & "' AND Type = 'E'", g_connData, adOpenDynamic, adLockOptimistic)
 
If m_bCopiePiece = False Then
 Do While Not rstItemFRS.EOF
 rstItemFRS.Fields("PIECE") = txtNoItem.Text
 
 Call rstItemFRS.Update
 
1  Call rstItemFRS.MoveNext
 Loop
 Else
 Set rstItemFRSDest = New ADODB.Recordset

 Call rstItemFRSDest.Open("SELECT * FROM GrbPiecesFRS", g_connData, adOpenDynamic, adLockOptimistic)

 Do While Not rstItemFRS.EOF
 Call rstItemFRSDest.AddNew

 rstItemFRSDest.Fields("IDFRS") = rstItemFRS.Fields("IDFRS")
 rstItemFRSDest.Fields("PIECE") = sNoPiece
 rstItemFRSDest.Fields("PRIX_SP") = rstItemFRS.Fields("PRIX_SP")
 rstItemFRSDest.Fields("PERS_RESS") = rstItemFRS.Fields("PERS_RESS")
 rstItemFRSDest.Fields("PRIX_LIST") = rstItemFRS.Fields("PRIX_LIST")
 rstItemFRSDest.Fields("ESCOMPTE") = rstItemFRS.Fields("ESCOMPTE")
 rstItemFRSDest.Fields("PRIX_NET") = rstItemFRS.Fields("PRIX_NET")
 rstItemFRSDest.Fields("DATE") = rstItemFRS.Fields("DATE")
 rstItemFRSDest.Fields("ENTRER_PAR") = rstItemFRS.Fields("ENTRER_PAR")
 rstItemFRSDest.Fields("VALIDE") = rstItemFRS.Fields("VALIDE")
 rstItemFRSDest.Fields("QUOTER") = rstItemFRS.Fields("QUOTER")
 rstItemFRSDest.Fields("DeviseMonétaire") = rstItemFRS.Fields("DeviseMonétaire")
 rstItemFRSDest.Fields("Type") = rstItemFRS.Fields("Type")

 Call rstItemFRSDest.Update

 Call rstItemFRS.MoveNext
3 Loop

 Call rstItemFRSDest.Close
 Set rstItemFRSDest = Nothing
End If

Call rstItemFRS.Close
Set rstItemFRS = Nothing

If m_bCopiePiece = True Then
 Call rstItem.AddNew
End If
 
 'Enregistrement des valeurs dans la table catalogue
rstItem.Fields("CATEGORIE") = txtCategorie.Text
rstItem.Fields("PIECE").Value = sNoPiece

3  For iCompteur = 1 To Len(sNoPiece)
 sLettre = Mid$(sNoPiece, iCompteur, 1)

If (Asc(sLettre) >= 4 And Asc(sLettre) <= 57) Or _
 (Asc(sLettre) >= 65 And Asc(sLettre) <= 90) Or _
 (Asc(sLettre) >=   And Asc(sLettre) <= 122) Then
 sPieceModif = sPieceModif & sLettre
End If
Next

3  rstItem.Fields("PIECE_MODIF").Value = sPieceModif
 rstItem.Fields("FABRICANT").Value = sNomFab
40 rstItem.Fields("PIECE_GRB").Value = txtNoItemGRB.Text
rstItem.Fields("DESC_EN").Value = txtDescriptionEN.Text
4 rstItem.Fields("DESC_FR").Value = txtDescriptionFR.Text
4 rstItem.Fields("TEMPS").Value = txtTemps.Text
4 rstItem.Fields("COMMENTAIRE").Value = txtComment.Text

4 Call rstItem.Update
 
4 Call rstItem.Close
4 Set rstItem = Nothing

4 If chkInventaire.Value = vbChecked Then
4 Set rstInventaire = New ADODB.Recordset

4 If m_bCopiePiece = True Then
4 Call rstInventaire.Open("SELECT * FROM GrbInventaireElec WHERE NoItem = '" & Replace(sNoPiece, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
4  Else
4  Call rstInventaire.Open("SELECT * FROM GrbInventaireElec WHERE NoItem = '" & Replace(m_sNoItem, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
4  End If

4  If rstInventaire.EOF Then
4  Call rstInventaire.AddNew
4  End If

4  rstInventaire.Fields("NoItem") = sNoPiece

4  rstInventaire.Fields("Description") = txtDescriptionFR.Text

50 rstInventaire.Fields("Manufacturier") = sNomFab

5 If chkBoite.Value = vbChecked Then
 rstInventaire.Fields("CommandeParBoite") = True
 rstInventaire.Fields("QteBoite") = txtQuantitéBoite.Text
 Else
 rstInventaire.Fields("CommandeParBoite") = False
 rstInventaire.Fields("QteBoite") = ""
 End If

 Set rstItemFRS = New ADODB.Recordset
 
 Call rstItemFRS.Open("SELECT * FROM GrbPiecesFRS WHERE PIECE = '" & Replace(sNoPiece, "'", "''") & "' AND IDFRS = 717", g_connData, adOpenDynamic, adLockOptimistic)
 
 If rstItemFRS.EOF Then
 Call rstItemFRS.AddNew

5  rstItemFRS.Fields("PIECE").Value = sNoPiece
5  rstItemFRS.Fields("IDFRS").Value = 717
5  rstItemFRS.Fields("Type").Value = "E"
5  rstItemFRS.Fields("PERS_RESS").Value = Null
5  rstItemFRS.Fields("PRIX_LIST").Value = "0"
5  rstItemFRS.Fields("ESCOMPTE").Value = "0"
5  rstItemFRS.Fields("PRIX_NET").Value = "0"
5  rstItemFRS.Fields("PrixReel").Value = "0"
60 rstItemFRS.Fields("DATE").Value = ConvertDate(Date)
  rstItemFRS.Fields("ENTRER_PAR").Value = g_sInitiale
  rstItemFRS.Fields("DeviseMonétaire").Value = "CAN"

  Call rstItemFRS.Update
  End If

  If chkBoite.Value = vbChecked Then
  If IsNumeric(rstItemFRS.Fields("PRIX_LIST")) Then
  rstInventaire.Fields("Prix Liste") = Round(rstItemFRS.Fields("PRIX_LIST") / txtQuantitéBoite.Text, 6)
  Else
  rstInventaire.Fields("Prix Liste") = "0"
  End If

  If IsNumeric(rstItemFRS.Fields("ESCOMPTE")) Then
6  rstInventaire.Fields("Escompte") = rstItemFRS.Fields("Escompte")
6  Else
6  rstInventaire.Fields("Escompte") = "0"
6  End If

6  If IsNumeric(rstItemFRS.Fields("PRIX_NET")) Then
6  rstInventaire.Fields("Prix net") = Round(rstItemFRS.Fields("PRIX_NET") / txtQuantitéBoite.Text, 6)
6  Else
6  rstInventaire.Fields("Prix net") = "0"
70 End If
  Else
  rstInventaire.Fields("Prix Liste") = rstItemFRS.Fields("PRIX_LIST")
  rstInventaire.Fields("Escompte") = rstItemFRS.Fields("Escompte")
  rstInventaire.Fields("Prix net") = rstItemFRS.Fields("PRIX_NET")
  End If

  Call rstItemFRS.Close
  Set rstItemFRS = Nothing

  rstInventaire.Fields("Commentaires") = txtComment.Text

  rstInventaire.Fields("Localisation") = cmbLocalisation.Text

  If Trim$(txtQuantiteStock.Text) <> "" Then
  rstInventaire.Fields("QuantitéStock") = txtQuantiteStock.Text
   Else
   rstInventaire.Fields("QuantitéStock") = "0"
7  End If

7  If chkMinimum.Value = vbChecked Then
7  rstInventaire.Fields("Minimum") = True

7  If Trim$(txtQuantiteMinimum.Text) <> "" Then
7  rstInventaire.Fields("QuantitéMinimum") = txtQuantiteMinimum.Text
7  Else
80 rstInventaire.Fields("QuantitéMinimum") = "0"
  End If

  If Trim$(txtQuantiteCommande.Text) = True Then
  rstInventaire.Fields("Commande") = txtQuantiteCommande.Text
  Else
  rstInventaire.Fields("Commande") = "0"
  End If
  Else
  rstInventaire.Fields("Minimum") = False
  rstInventaire.Fields("QuantitéMinimum") = ""
  rstInventaire.Fields("Commande") = ""
  End If

   Call rstInventaire.Update

   Call rstInventaire.Close
   Set rstInventaire = Nothing
   Else
8  If m_bCopiePiece = True Then
8  Call g_connData.Execute("DELETE * FROM GrbInventaireElec WHERE NoItem = '" & Replace(sNoPiece, "'", "''") & "'")
8  Else
8  Call g_connData.Execute("DELETE * FROM GrbInventaireElec WHERE NoItem = '" & Replace(m_sNoItem, "'", "''") & "'")
90 End If
90 End If

  If m_bCopiePiece = False Then
  If txtQuantiteStock.Text <> m_sQteStockAvant Or ((m_sQteStockAvant <> "" And m_sQteStockAvant <> "0") And chkInventaire.Value = vbUnchecked) Then
  If m_sQteStockAvant = "" Then
  m_sQteStockAvant = "0"
  End If

  If Not IsNumeric(txtQuantiteStock.Text) Then
  txtQuantiteStock.Text = "0"
  End If

  Set rstInvModif = New ADODB.Recordset

  Call rstInvModif.Open("SELECT * FROM GrbInventaireElecModif", g_connData, adOpenDynamic, adLockOptimistic)

 Call rstInvModif.AddNew

   rstInvModif.Fields("Date") = ConvertDate(Date)
 rstInvModif.Fields("IDProjet") = InputBox("Précisez l'ajustement d'inventaire")
   rstInvModif.Fields("NoItem") = txtNoItem.Text

 If chkInventaire.Value = vbChecked Then
   rstInvModif.Fields("Quantité") = CDbl(txtQuantiteStock.Text) - CDbl(m_sQteStockAvant)
 Else
9  rstInvModif.Fields("Quantité") = 0 - CDbl(m_sQteStockAvant)
 End If

10 rstInvModif.Fields("User") = g_sInitiale

1 Call rstInvModif.Update

1 Call rstInvModif.Close
 Set rstInvModif = Nothing
1End If
End If

10 If (UCase(sNoPiece) <> UCase(m_sNoItem)) And m_bCopiePiece = False Then
 Call ModifierNoItem(m_sNoItem, sNoPiece)
10 End If

m_sQteStockAvant = ""

10 m_bRempliManuel = True
 
10  m_sSelectNoItem = sNoPiece
10  m_sSelectFabricant = sNomFab

10  Call RemplirComboLocalisation
 
10  Call RemplirComboFabricant
 
 'Rétablir les buttons
10  Call MontrerControles(MODE_INACTIF)
 
10  Call BarrerChamps_piece(True)

109Exit Sub

Oups:

10  wOups "frmCatalogueElec", "EnregistrerItem", Err, Err.number, Err.Description
End Sub

Private Sub cmdChangerCategorie_Click()

 On Error GoTo Oups

 Dim rstPiece As ADODB.Recordset
 
 Call frmChoixCategorie.Afficher(ELECTRIQUE)
 
 If txtCategorie.Text <> m_sCategorieCopie Then
 If m_bAnnulerCopie = False Then
 Set rstPiece = New ADODB.Recordset

 Call rstPiece.Open("SELECT * FROM GrbCatalogueElec WHERE PIECE = '" & Replace(cmbNoItem.Text, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)

 rstPiece.Fields("CATEGORIE") = m_sCategorieCopie

 Call rstPiece.Update

 Call rstPiece.Close
 Set rstPiece = Nothing
 
  Call ViderChamps_piece

  m_sSelectFabricant = txtFabricant.Text

  Call RemplirComboFabricant

  Call MontrerControles(MODE_INACTIF)

  Call BarrerChamps_piece(True)
  End If
  End If

  Exit Sub

Oups:

10 wOups "frmCatalogueElec", "cmdChangerCategorie_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdCopier_Click()
 
 On Error GoTo Oups

 m_bCopiePiece = True

 Call CmdModif_Click

 chkInventaire.Value = vbUnchecked
 chkBoite.Value = vbUnchecked
 chkMinimum.Value = vbUnchecked

 txtQuantitéBoite.Text = ""
 txtQuantiteStock.Text = ""
 txtQuantiteMinimum.Text = ""
 txtQuantiteCommande.Text = ""
 cmbLocalisation.Text = ""

  Exit Sub

Oups:
 
  wOups "frmCatalogueElec", "cmdCopier_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdDemande_Click()

 On Error GoTo Oups

 Call dlgDemandePrix.Afficher(Me)
 
 If m_bDemandeAnnuler = False Then
 Call frmChoixDemande.Afficher(ELECTRIQUE, m_eDemande)
 End If

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "cmdDemande_Click", Err, Err.number, Err.Description
End Sub

Private Sub CmdEnr_Click()

 On Error GoTo Oups

 Dim iCompteur As Integer
 Dim bContinuer As Boolean

 'Enregistrement d'un item dans la BD
 txtFabricant.Top = 1320 'GLL 2017-10-10 Modification pour ajouter la modification possible du Manufacturier
 cmbFabricant.Visible = True 'GLL 2017-10-10 Modification pour ajouter la modification possible du Manufacturier
 If (UCase(txtNoItem.Text) <> UCase(m_sNoItem)) And m_bCopiePiece = False Then
 If MsgBox("Le numéro de pièce sera modifié dans toutes les soumissions, les projets et les achats. " & vbNewLine & _
 "Voulez-vous continuer ? ", vbYesNo) = vbYes Then
 bContinuer = True
 Else
 bContinuer = False
 End If
 Else
 bContinuer = True
  End If
 
  If bContinuer = True Then
  Call EnregistrerItem

  If m_eMode = MODE_INACTIF Then
  m_bCopiePiece = False
  End If

  Call RemplirComboDescription

  m_bBloqueDescription = True

For iCompteur = 0 To cmbDescriptionFR.ListCount - 1
If cmbDescriptionFR.LIST(iCompteur) = txtDescriptionFR.Text Then
 cmbDescriptionFR.ListIndex = iCompteur

 Exit For
 End If
 Next

 m_bBloqueDescription = False
End If

Exit Sub

Oups:

wOups "frmCatalogueElec", "CmdEnr_Click", Err, Err.number, Err.Description
End Sub

Private Sub ModifierNoItem(ByVal sAncienNoItem As String, ByVal sNouveauNoItem As String)
 
 On Error GoTo Oups

 Dim iRecordProjet As Integer
 Dim iRecordSoum As Integer
 Dim iRecordAchat As Integer

 Call g_connData.Execute("UPDATE GrbProjet_Pieces SET NumItem = '" & Replace(sNouveauNoItem, "'", "''") & "' WHERE NumItem = '" & Replace(sAncienNoItem, "'", "''") & "' AND Type = 'E'", iRecordProjet)
 Call g_connData.Execute("UPDATE GrbSoumission_Pieces SET NumItem = '" & Replace(sNouveauNoItem, "'", "''") & "' WHERE NumItem = '" & Replace(sAncienNoItem, "'", "''") & "' AND Type = 'E'", iRecordSoum)

 Call g_connData.Execute("UPDATE GrbAchat_Pieces SET PIECE = '" & Replace(sNouveauNoItem, "'", "''") & "' WHERE PIECE = '" & Replace(sAncienNoItem, "'", "''") & "' AND Left(IDAchat, 1) <> 'M'", iRecordAchat)

 Call g_connData.Execute("UPDATE GrbInventaireElecModif SET NoItem = '" & Replace(sNouveauNoItem, "'", "''") & "' WHERE NoItem = '" & Replace(sAncienNoItem, "'", "''") & "'")

 Call MsgBox("Numéros de pièces modifiés" & vbNewLine & _
 vbNewLine & _
 "Projets : " & iRecordProjet & vbNewLine & _
 "Soumissions : " & iRecordSoum & vbNewLine & _
 "Achats : " & iRecordAchat, vbOKOnly)

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "ModifierNoItem", Err, Err.number, Err.Description
End Sub

Private Sub cmdEnrFrs_Click()

 On Error GoTo Oups

 Dim iCompteur As Integer
 'Enregistrement d'un Item dans la BD
 
 If Trim$(txtPrixList.Text) = "" And Trim$(mskEscompte.Text) = "" And Trim$(txtPrixNet.Text) = "" And Trim$(txtPrixSpecial.Text) = "" Then
 txtPrixList.Text = "0"
 mskEscompte.Text = "0"
 txtPrixNet.Text = "0"
 End If
 
 If Trim$(txtPrixList.Text) = vbNullString Then
 If Trim$(txtPrixSpecial.Text) = vbNullString Then
 Call MsgBox("Vous devez remplir le prix listé!", vbOKOnly, "Erreur")

 Exit Sub
  End If
  End If
 
  If Trim$(txtPrixNet.Text) = vbNullString And Trim$(txtPrixSpecial.Text) = vbNullString Then
  Call MsgBox("Vous devez remplir le prix net ou le prix spécial!", vbOKOnly, "Erreur")
 
  Exit Sub
  End If

  If optUSA.Value = True Or optSpain.Value = True Then
  If Trim$(txtTauxChange.Text) <> vbNullString Then
 If Not IsNumeric(txtTauxChange.Text) Then
 Call MsgBox("Le taux de change doit être numérique!", vbOKOnly, "Erreur")

 Exit Sub
 End If
 Else
 Call MsgBox("Le taux de change ne doit pas être vide!", vbOKOnly, "Erreur")

 Exit Sub
 End If
End If

 If (Trim$(txtPrixNet.Text) <> Trim$(txtPrixList.Text)) And Trim$(txtPrixSpecial.Text) = vbNullString Then
 Call CalculerPrixNet
End If

1  If cmbfrs.ListIndex > -1 Then
 Call EnregistrerFRS
 Else
 Call MsgBox("Vous devez choisir un fournisseur!", vbOKOnly, "Erreur")
 End If

Exit Sub

Oups:

 wOups "frmCatalogueElec", "cmdEnrFrs_Click", Err, Err.number, Err.Description
End Sub

Private Sub EnregistrerFRS()

 On Error GoTo Oups

 'Enregistrement de l'item dans la BD
 Dim rstItemFRS As ADODB.Recordset
 Dim rstInv As ADODB.Recordset
 Dim rstConfig As ADODB.Recordset
 Dim bDistribExiste As Boolean
 Dim iCompteur As Integer
 
 'Si le PRIX_SP est monétaire
 If txtPrixSpecial.Text <> vbNullString Then
 If Not IsNumeric(txtPrixSpecial.Text) Then
 Call MsgBox("Le prix spécial est invalide!", vbOKOnly, "Erreur")
 
 Exit Sub
 End If
  End If
 
 'Si le PRIX_NET est monétaire
  If txtPrixNet.Text <> vbNullString Then
  If Not IsNumeric(txtPrixNet.Text) Then
  Call MsgBox("Le prix net est invalide!", vbOKOnly, "Erreur")
 
  Exit Sub
  End If
  End If
 
 'Si le PRIX_LIST est monétaire
  If txtPrixList.Text <> vbNullString Then
If Not IsNumeric(txtPrixList.Text) Then
Call MsgBox("Le prix listé est invalide!", vbOKOnly, "Erreur")
 
 Exit Sub
 End If
End If
 
 'Si la date de validité est valide
If Trim$(mskValide.Text) <> vbNullString Then
 If IsDate(mskValide.Text) = False Then
 Call MsgBox("La date de validité est invalide!", vbOKOnly, "Erreur")
 
 Exit Sub
 End If
End If

bDistribExiste = False
 
1  If m_bAjout = True Then
 'Si le distributeur n'est pas déjà dans le listView
 If lvwfournisseur.ListItems.count > 0 Then
 For iCompteur = 1 To lvwfournisseur.ListItems.count
 If lvwfournisseur.ListItems(iCompteur).Text = cmbfrs.Text Then
 bDistribExiste = True
 
 Exit For
 End If
1  Next
 End If
 
 If bDistribExiste = True Then
 If txtPrixSpecial.Text <> "" Then
 If lvwfournisseur.ListItems(iCompteur).SubItems(I_COL_FRS_PRIX_SP) <> "" Then
 Call MsgBox("Ce distributeur est déjà ajouté avec un prix spécial", vbOKOnly, "Erreur")

 Exit Sub
 End If
 Else
 If lvwfournisseur.ListItems(iCompteur).SubItems(I_COL_FRS_PRIX_NET) <> "" Then
 Call MsgBox("Ce distributeur est déjà ajouté avec un prix net", vbOKOnly, "Erreur")

 Exit Sub
 End If
 End If
 End If
2  End If

Set rstItemFRS = New ADODB.Recordset

2  If m_bAjout = True Then
 Call rstItemFRS.Open("SELECT * FROM GrbPiecesFRS WHERE PIECE = '" & Replace(txtNoItem.Text, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 'si c'est un ajout, j'ouvre un recordset général
Call rstItemFRS.AddNew
 
 m_bAjout = False
30 Else
3 Call rstItemFRS.Open("SELECT * FROM GrbPiecesFRS WHERE noEnreg = " & lvwfournisseur.SelectedItem.Tag, g_connData, adOpenDynamic, adLockOptimistic)
End If
 
 'Enregistrement des valeurs dans la table catalogue
rstItemFRS.Fields("PIECE").Value = cmbNoItem.Text
rstItemFRS.Fields("IDFRS").Value = cmbfrs.ItemData(cmbfrs.ListIndex)
rstItemFRS.Fields("Type").Value = "E"
 
If cmbPersRess.ListIndex > -1 Then
 rstItemFRS.Fields("PERS_RESS").Value = cmbPersRess.ItemData(cmbPersRess.ListIndex)
Else
 rstItemFRS.Fields("PERS_RESS").Value = Null
End If
 
rstItemFRS.Fields("PRIX_LIST").Value = txtPrixList.Text
3  rstItemFRS.Fields("ESCOMPTE").Value = mskEscompte.Text
 
If txtPrixSpecial.Text <> vbNullString Or txtPrixNet.Text <> vbNullString Then
If txtPrixNet.Text <> vbNullString Then
 rstItemFRS.Fields("PRIX_NET").Value = txtPrixNet.Text
 rstItemFRS.Fields("PrixReel").Value = txtPrixNet.Text
 Else
 rstItemFRS.Fields("PRIX_NET").Value = vbNullString
 End If

If txtPrixSpecial.Text <> vbNullString Then
4 rstItemFRS.Fields("PRIX_SP").Value = txtPrixSpecial.Text
4 rstItemFRS.Fields("PrixReel").Value = txtPrixNet.Text
4 Else
4 rstItemFRS.Fields("PRIX_SP").Value = vbNullString
4 End If
4 End If
 
4 rstItemFRS.Fields("DATE").Value = ConvertDate(Date)
4 rstItemFRS.Fields("VALIDE").Value = mskValide.Text
4 rstItemFRS.Fields("ENTRER_PAR").Value = g_sInitiale
 
4 If chkquoter.Value = 1 Then
4 rstItemFRS.Fields("quoter").Value = True
4  Else
4  rstItemFRS.Fields("quoter").Value = False
4  End If

4  If optCAN.Value = True Then
4  rstItemFRS.Fields("devisemonétaire").Value = "CAN"
4  Else
4  If optUSA.Value = True Then
4  rstItemFRS.Fields("DeviseMonétaire").Value = "USA"
50 Else
rstItemFRS.Fields("DeviseMonétaire").Value = "SPA"
 End If
 End If
 
 Call rstItemFRS.Update
 
 Call rstItemFRS.Close
 Set rstItemFRS = Nothing
 
 If optUSA.Value = True Or optSpain.Value = True Then
 Set rstConfig = New ADODB.Recordset

 Call rstConfig.Open("SELECT TauxAmericain, TauxEspagnol FROM GrbConfig", g_connData, adOpenDynamic, adLockOptimistic)

 If optUSA.Value = True Then
 rstConfig.Fields("TauxAmericain") = txtTauxChange.Text
5  Else
5  rstConfig.Fields("TauxEspagnol") = txtTauxChange.Text
5  End If

5  Call rstConfig.Update

5  Call rstConfig.Close
5  Set rstConfig = Nothing
5  End If

 'Rétablir les boutons
5  Call MontrerControles(MODE_INACTIF)

60 If cmbfrs.ItemData(cmbfrs.ListIndex) = 71 Then
  Set rstInv = New ADODB.Recordset

  Call rstInv.Open("SELECT * FROM GrbInventaireElec WHERE NoItem = '" & Replace(txtNoItem.Text, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)

  If Not rstInv.EOF Then
  If txtPrixNet.Text <> "" Then
  If rstInv.Fields("CommandeParBoite") = True Then
  rstInv.Fields("Prix Liste") = txtPrixList.Text / rstInv.Fields("QteBoite")
  rstInv.Fields("Escompte") = mskEscompte.Text
  rstInv.Fields("Prix net") = txtPrixNet.Text / rstInv.Fields("QteBoite")
  Else
  rstInv.Fields("Prix Liste") = txtPrixList.Text
  rstInv.Fields("Escompte") = mskEscompte.Text
6  rstInv.Fields("Prix net") = txtPrixNet.Text
6  End If
6  Else
6  If rstInv.Fields("CommandeParBoite") = True Then
6  rstInv.Fields("Prix Liste") = txtPrixSpecial.Text / rstInv.Fields("QteBoite")
6  rstInv.Fields("Escompte") = ""
6  rstInv.Fields("Prix net") = txtPrixSpecial.Text / rstInv.Fields("QteBoite")
6  Else
70 rstInv.Fields("Prix Liste") = txtPrixSpecial.Text
  rstInv.Fields("Escompte") = ""
  rstInv.Fields("Prix net") = txtPrixSpecial.Text
  End If
  End If

  Call rstInv.Update
  End If

  Call rstInv.Close
  Set rstInv = Nothing
  End If
 
 'Remplis le ListView
  Call RemplirListViewFournisseur

  Exit Sub

Oups:

   wOups "frmCatalogueElec", "EnregistrerFRS", Err, Err.number, Err.Description
End Sub

Private Sub CmdFerme_Click()

 On Error GoTo Oups
 
 'Fermeture de la fenêtre
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "CmdFerme_Click", Err, Err.number, Err.Description
End Sub

Private Sub CmdModif_Click()

 On Error GoTo Oups

 'procedure qui permet de modifier l'enregistrement courant
 'on montre/cache les maskedBox
 If cmbNoItem.ListCount > 0 Then
 
 'Copie le contenu textbox dans les maskbox
 Call MontrerControles(MODE_AJOUT_MODIF_ELEC)
 txtFabricant.Top = 960 'GLL 2017-10-10 Modification pour ajouter la modification possible du Manufacturier
 cmbFabricant.Visible = False 'GLL 2017-10-10 Modification pour ajouter la modification possible du Manufacturier
 
 m_sQteStockAvant = txtQuantiteStock.Text
 
 Call BarrerChamps_piece(False)
 End If

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "CmdModif_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdModifFrs_Click()

 On Error GoTo Oups

 'modifie un fournisseur pour la piece
 If lvwfournisseur.ListItems.count > 0 Then
 Call ModifierFournisseur
 End If

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "cmdModifFrs_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdRechercheCategorie_Click()
 On Error GoTo Oups

 Dim rstcatalog As ADODB.Recordset
 Dim sDescription As String
 Dim itmDescription As ListItem
 'ouvre un boite de dialogue pour savoir quoi rechercher
 sDescription = InputBox("Quelle est la description à rechercher")
 
 If sDescription <> vbNullString Then 'Si il y a quelque chose a chercher
 Call lvwCategorie.ListItems.Clear 'Vide la liste pour ne pas avoir l'ancienne recherche
 
 sDescription = Replace(sDescription, "'", "''")
 
 sDescription = "%" & sDescription & "%"

 Set rstcatalog = New ADODB.Recordset

55
 Call rstcatalog.Open("SELECT DISTINCT CATEGORIE FROM GrbCatalogueElec WHERE Categorie LIKE '" & sDescription & "' ORDER BY CATEGORIE", g_connData, adOpenDynamic, adLockOptimistic)
 'Rempli la liste pour pouvoir le sélectionner
  Do While Not rstcatalog.EOF
  Set itmDescription = lvwCategorie.ListItems.Add()
 
  itmDescription.Tag = rstcatalog.Fields("CATEGORIE")
 itmDescription.Text = rstcatalog.Fields("CATEGORIE")

 Call rstcatalog.MoveNext
Loop
 'Fermeture de la table
 Call rstcatalog.Close
 Set rstcatalog = Nothing
 'si il y a des choix posible on les affiche
 If lvwCategorie.ListItems.count > 0 Then
 lvwCategorie.Visible = True

 Call lvwCategorie.SetFocus
 Else
1  Call MsgBox("Aucun enregistrement trouvé!")
 End If
 End If

Exit Sub

Oups:

wOups "frmCatalogueElec", "cmdRechercheDescrFR_Click", Err, Err.number, Err.Description

End Sub

Private Sub cmdRechercheInventaire_Click()

 On Error GoTo Oups

 Call frmRechercheInventaire.Afficher(ELECTRIQUE)

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "cmdRechercheInventaire_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdRechercheAchat_Click()

 On Error GoTo Oups

 Dim rstAchat As ADODB.Recordset
 Dim itmAchat As ListItem

 Screen.MousePointer = vbHourglass

 Call lvwRechercheAchat.ListItems.Clear

 Set rstAchat = New ADODB.Recordset

 Call rstAchat.Open("SELECT DISTINCT (IDAchat + '-' + RIGHT('00' & IndexAchat,3)) As NumeroAchat, SUM(Qté) As QtéTotale FROM GrbAchat_Pieces WHERE TRIM(PIECE) = '" & Trim$(Replace(txtNoItem.Text, "'", "''")) & "' GROUP BY (IDAchat + '-' + RIGHT('00' & IndexAchat,3))", g_connData, adOpenForwardOnly, adLockReadOnly)

 Do While Not rstAchat.EOF
 Set itmAchat = lvwRechercheAchat.ListItems.Add

 itmAchat.Text = rstAchat.Fields("NumeroAchat")
 itmAchat.SubItems(I_COL_ACHAT_QUANTITE) = rstAchat.Fields("QtéTotale")

  Call rstAchat.MoveNext
  Loop

  Call rstAchat.Close
  Set rstAchat = Nothing

  Screen.MousePointer = vbDefault

  If lvwRechercheAchat.ListItems.count > 0 Then
  lvwRechercheAchat.Visible = True

  Call lvwRechercheAchat.SetFocus
10 Else
1 Call MsgBox("Cette pièce n'a jamais été utilisée dans les achats!", vbOKOnly)
End If

Exit Sub

Oups:

wOups "frmCatalogueElec", "cmdRechercheAchat_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdRechercheJob_Click()

 On Error GoTo Oups

 Dim rstProjSoum As ADODB.Recordset
 Dim itmProjSoum As ListItem

 Screen.MousePointer = vbHourglass

 Call lvwRechercheJob.ListItems.Clear

 Set rstProjSoum = New ADODB.Recordset

 Call rstProjSoum.Open("SELECT DISTINCT IDProjet, SUM(Qté) As QtéTotale FROM GrbProjet_Pieces WHERE TRIM(NumItem) = '" & Trim$(Replace(txtNoItem.Text, "'", "''")) & "' And Type = 'E' GROUP BY IDProjet", g_connData, adOpenForwardOnly, adLockReadOnly)

 Do While Not rstProjSoum.EOF
 Set itmProjSoum = lvwRechercheJob.ListItems.Add

 itmProjSoum.Text = rstProjSoum.Fields("IDProjet")
 itmProjSoum.SubItems(I_COL_JOB_QUANTITE) = rstProjSoum.Fields("QtéTotale")

  Call rstProjSoum.MoveNext
  Loop

  Call rstProjSoum.Close

  Call rstProjSoum.Open("SELECT DISTINCT IDSoumission, SUM(Qté) As QtéTotale FROM GrbSoumission_Pieces WHERE TRIM(NumItem) = '" & Trim$(Replace(txtNoItem.Text, "'", "''")) & "' And Type = 'E' GROUP BY IDSoumission", g_connData, adOpenForwardOnly, adLockReadOnly)

  Do While Not rstProjSoum.EOF
  Set itmProjSoum = lvwRechercheJob.ListItems.Add

  itmProjSoum.Text = rstProjSoum.Fields("IDSoumission")
  itmProjSoum.SubItems(I_COL_JOB_QUANTITE) = rstProjSoum.Fields("QtéTotale")

Call rstProjSoum.MoveNext
Loop

Call rstProjSoum.Close
Set rstProjSoum = Nothing

Screen.MousePointer = vbDefault

If lvwRechercheJob.ListItems.count > 0 Then
 lvwRechercheJob.Visible = True

 Call lvwRechercheJob.SetFocus
Else
 Call MsgBox("Cette pièce n'a jamais été utilisée dans les jobs!", vbOKOnly)
End If

Exit Sub

Oups:

1  wOups "frmCatalogueElec", "cmdRechercheJob_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdRechercherManufacturier_Click()

 On Error GoTo Oups

 Dim rstManufact As ADODB.Recordset
 Dim sManufact As String
 Dim itmManufact As ListItem
 
 sManufact = InputBox("Quel est le manufacturier à rechercher?")
 
 sManufact = Replace(sManufact, "'", "''")
 
 If sManufact <> vbNullString Then
 Call lvwFabricant.ListItems.Clear
 
 Set rstManufact = New ADODB.Recordset

 Call rstManufact.Open("SELECT * FROM GrbCatalogueElec WHERE INSTR(1, FABRICANT, '" & sManufact & "') > 0 ORDER BY FABRICANT", g_connData, adOpenDynamic, adLockOptimistic)

 Do While Not rstManufact.EOF
  Set itmManufact = lvwFabricant.ListItems.Add
 
  itmManufact.Tag = rstManufact.Fields("CATEGORIE")
 
  itmManufact.Text = Trim(rstManufact.Fields("FABRICANT"))

  itmManufact.SubItems(I_COL_MAN_PIECE) = Trim(rstManufact.Fields("PIECE"))
 
  If Not IsNull(rstManufact.Fields("DESC_FR")) Then
  itmManufact.SubItems(I_COL_MAN_DESCR_FR) = Trim(rstManufact.Fields("DESC_FR"))
  Else
  itmManufact.SubItems(I_COL_MAN_DESCR_FR) = vbNullString
 End If
 
If Not IsNull(rstManufact.Fields("DESC_EN")) Then
 itmManufact.SubItems(I_COL_MAN_DESCR_EN) = Trim(rstManufact.Fields("DESC_EN"))
 Else
 itmManufact.SubItems(I_COL_MAN_DESCR_EN) = vbNullString
 End If
 
 Call rstManufact.MoveNext
 Loop
 
 Call rstManufact.Close
 Set rstManufact = Nothing
 
 If lvwFabricant.ListItems.count > 0 Then
 lvwFabricant.Visible = True
 
 Call lvwFabricant.SetFocus
 Else
 Call MsgBox("Aucun enregistrement trouvé!")
 End If
 End If

Exit Sub

Oups:

 wOups "frmCatalogueElec", "cmdRechercherManufacturier_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdRechercherPiece_Click()

 On Error GoTo Oups

 Dim rstPiece As ADODB.Recordset
 Dim sPiece As String
 Dim itmPiece As ListItem
 Dim iCompteur As Integer
 Dim sPieceModif As String
 Dim sLettre As String
 
 sPiece = InputBox("Quelle est la pièce à rechercher?")
 
 If sPiece <> vbNullString Then
 For iCompteur = 1 To Len(sPiece)
 sLettre = Mid$(sPiece, iCompteur, 1)
 
  If (Asc(sLettre) >= 4 And Asc(sLettre) <= 57) Or _
 (Asc(sLettre) >= 65 And Asc(sLettre) <= 90) Or _
 (Asc(sLettre) >=   And Asc(sLettre) <= 122) Then
  sPieceModif = sPieceModif & sLettre
  End If
  Next

  Call lvwPieces.ListItems.Clear

  Set rstPiece = New ADODB.Recordset

  Call rstPiece.Open("SELECT * FROM GrbCatalogueElec WHERE INSTR(1, PIECE_MODIF, '" & sPieceModif & "') > 0 ORDER BY PIECE", g_connData, adOpenDynamic, adLockOptimistic)

  Do While Not rstPiece.EOF
 Set itmPiece = lvwPieces.ListItems.Add
 
itmPiece.Text = Trim(rstPiece.Fields("PIECE"))

 If Not IsNull(rstPiece.Fields("FABRICANT")) Then
 itmPiece.SubItems(I_COL_PIECE_FABRICANT) = Trim(rstPiece.Fields("FABRICANT"))
 Else
 itmPiece.SubItems(I_COL_PIECE_FABRICANT) = vbNullString
 End If
 
 If Not IsNull(rstPiece.Fields("DESC_FR")) Then
 itmPiece.SubItems(I_COL_PIECE_DESCR_FR) = Trim(rstPiece.Fields("DESC_FR"))
 Else
 itmPiece.SubItems(I_COL_PIECE_DESCR_FR) = vbNullString
 End If
 
 If Not IsNull(rstPiece.Fields("DESC_EN")) Then
 itmPiece.SubItems(I_COL_PIECE_DESCR_EN) = Trim(rstPiece.Fields("DESC_EN"))
 Else
 itmPiece.SubItems(I_COL_PIECE_DESCR_EN) = vbNullString
 End If
 
 itmPiece.Tag = rstPiece.Fields("CATEGORIE")
 
 Call rstPiece.MoveNext
1  Loop
 
 Call rstPiece.Close
 Set rstPiece = Nothing
 
 If lvwPieces.ListItems.count > 0 Then
 lvwPieces.Visible = True
 
 Call lvwPieces.SetFocus
 Else
 Call MsgBox("Aucun enregistrement trouvé!")
 End If
End If

Exit Sub

Oups:

wOups "frmCatalogueElec", "cmdRechercherPiece_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdRechercheDescrFR_Click()

 On Error GoTo Oups

 Dim rstDescription As ADODB.Recordset
 Dim sDescription As String
 Dim itmDescription As ListItem
 
 sDescription = InputBox("Quelle est la description à rechercher")
 
 If sDescription <> vbNullString Then
 Call lvwDescription.ListItems.Clear
 
 sDescription = Replace(sDescription, "'", "''")
 
 sDescription = "%" & sDescription & "%"

 Set rstDescription = New ADODB.Recordset

 Call rstDescription.Open("SELECT * FROM GrbCatalogueElec WHERE DESC_FR LIKE '" & sDescription & "' OR DESC_EN LIKE '" & sDescription & "' ORDER BY DESC_FR, DESC_EN", g_connData, adOpenDynamic, adLockOptimistic)
 
  Do While Not rstDescription.EOF
  Set itmDescription = lvwDescription.ListItems.Add()
 
  itmDescription.Tag = rstDescription.Fields("CATEGORIE")
 
  If Not IsNull(rstDescription.Fields("DESC_FR")) Then
  itmDescription.Text = Trim(rstDescription.Fields("DESC_FR"))
  Else
  itmDescription.Text = vbNullString
  End If
 
 If Not IsNull(rstDescription.Fields("DESC_EN")) Then
 itmDescription.SubItems(I_COL_DES_DESCR_EN) = Trim(rstDescription.Fields("DESC_EN"))
 Else
 itmDescription.SubItems(I_COL_DES_DESCR_EN) = vbNullString
 End If
 
 If Not IsNull(rstDescription.Fields("FABRICANT")) Then
 itmDescription.SubItems(I_COL_DES_FABRICANT) = Trim(rstDescription.Fields("FABRICANT"))
 Else
 itmDescription.SubItems(I_COL_DES_FABRICANT) = vbNullString
 End If

 itmDescription.SubItems(I_COL_DES_PIECE) = Trim(rstDescription.Fields("PIECE"))
 
 Call rstDescription.MoveNext
Loop
 
 Call rstDescription.Close
 Set rstDescription = Nothing

 If lvwDescription.ListItems.count > 0 Then
 lvwDescription.Visible = True

 Call lvwDescription.SetFocus
 Else
1  Call MsgBox("Aucun enregistrement trouvé!")
 End If
 End If

Exit Sub

Oups:

wOups "frmCatalogueElec", "cmdRechercheDescrFR_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdTotal_Click()

 On Error GoTo Oups

 Dim sAnnee As String
 Dim rstTotal As ADODB.Recordset
 Dim dblTotalProj As Double
 Dim dblTotalAchat As Double

 sAnnee = InputBox("Pour quelle année? (AAAA)")

 If Len(sAnnee) = 4 Then
 If IsNumeric(sAnnee) Then
 If CInt(sAnnee) <= Year(Date) Then
 Screen.MousePointer = vbHourglass

 Set rstTotal = New ADODB.Recordset

  Call rstTotal.Open("SELECT SUM(Qté) As Total FROM GrbProjet_Pieces INNER JOIN GrbProjetElec ON GrbProjet_Pieces.IDProjet = GrbProjetElec.IDProjet WHERE TRIM(NumItem) = '" & Trim$(Replace(txtNoItem.Text, "'", "''")) & "' AND Type = 'E' AND Left(Creer,4) = '" & sAnnee & "' AND RIGHT(GrbProjet_Pieces.IDProjet,2) < '60'", g_connData, adOpenDynamic, adLockOptimistic)

  If Not IsNull(rstTotal.Fields("Total")) Then
  dblTotalProj = CDbl(rstTotal.Fields("Total"))
  Else
  dblTotalProj = 0
  End If

  Call rstTotal.Close

  Call rstTotal.Open("SELECT SUM(Qté) As Total FROM GrbAchat_Pieces INNER JOIN GrbAchat ON GrbAchat_Pieces.IDAchat = GrbAchat.IDAchat AND GrbAchat_Pieces.IndexAchat = GrbAchat.IndexAchat WHERE TRIM(PIECE) = '" & Trim$(Replace(txtNoItem.Text, "'", "''")) & "' AND Left(DateAchat,4) = '" & sAnnee & "'", g_connData, adOpenDynamic, adLockOptimistic)

 If Not IsNull(rstTotal.Fields("Total")) Then
 dblTotalAchat = CDbl(rstTotal.Fields("Total"))
 Else
 dblTotalAchat = 0
 End If

 Call rstTotal.Close
 Set rstTotal = Nothing

 Screen.MousePointer = vbDefault

 Call MsgBox("Quantité utilisée en " & sAnnee & " : " & vbNewLine & _
 vbNewLine & _
 "Projets : " & dblTotalProj & vbNewLine & _
 "Achats : " & dblTotalAchat)
 Else
 Call MsgBox("Année trop grande!", vbOKOnly, "Erreur")
 End If
Else
 Call MsgBox("Année non numérique!", vbOKOnly, "Erreur")
 End If
Else
 If Len(sAnnee) <> 0 Then
 Call MsgBox("L'année doit être sur 4 chiffres!", vbOKOnly, "Erreur")
 End If
1  End If

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "cmdTotal_Click", Err, Err.number, Err.Description
End Sub



Private Sub Form_Click()

 On Error GoTo Oups

 lvwDescription.Visible = False
 lvwFabricant.Visible = False
 lvwPieces.Visible = False
 lvwRechercheJob.Visible = False
 lvwRechercheAchat.Visible = False

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "Form_Click", Err, Err.number, Err.Description
End Sub

Private Sub fraApprob_Click()

 On Error GoTo Oups

 lvwDescription.Visible = False
 lvwFabricant.Visible = False
 lvwPieces.Visible = False
 lvwRechercheJob.Visible = False
 lvwRechercheAchat.Visible = False

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "fraApprob_Click", Err, Err.number, Err.Description
End Sub

Private Sub frafournisseur_Click()

 On Error GoTo Oups

 lvwDescription.Visible = False
 lvwFabricant.Visible = False
 lvwPieces.Visible = False
 lvwRechercheJob.Visible = False
 lvwRechercheAchat.Visible = False

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "fraFournisseur_Click", Err, Err.number, Err.Description
End Sub

Private Sub fraQuantité_Click()

 On Error GoTo Oups

 lvwDescription.Visible = False
 lvwFabricant.Visible = False
 lvwPieces.Visible = False
 lvwRechercheJob.Visible = False
 lvwRechercheAchat.Visible = False

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "fraQuantité_Click", Err, Err.number, Err.Description
End Sub



Private Sub lvwCategorie_DblClick()
5
 Dim itmDescription As ListItem
 Dim iCompteur As Integer

 If lvwCategorie.ListItems.count > 0 Then
 Screen.MousePointer = vbHourglass

 Set itmDescription = lvwCategorie.SelectedItem

 'm_sSelectCategorie = itmDescription.Tag
 'm_sSelectFabricant = Trim$(itmDescription.SubItems(I_COL_DES_FABRICANT))
 ' m_sSelectNoItem = Trim$(itmDescription.SubItems(I_COL_DES_PIECE))

 'If m_eMode = MODE_INACTIF Then
 ' Call RemplirComboCategorie
  'Else
  cmbCategorie.Text = lvwCategorie.SelectedItem.Text
  'pour pouvoir
 Call cmbCategorie_Click
  lvwCategorie.Visible = False

  Screen.MousePointer = vbDefault
  End If

  Exit Sub
End Sub

Private Sub lvwCategorie_LostFocus()
 On Error GoTo Oups

 If lvwCategorie.Visible = True Then
 lvwCategorie.Visible = False
 End If

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "lvwCategorie_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub lvwDescription_KeyDown(KeyCode As Integer, Shift As Integer)
 
 On Error GoTo Oups

 If KeyCode = vbKeyReturn Then
 Call lvwDescription_DblClick
 End If

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "lvwDescription_KeyDown", Err, Err.number, Err.Description
End Sub

Private Sub lvwFabricant_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

 On Error GoTo Oups

 lvwFabricant.Sorted = True
 
 If lvwFabricant.SortOrder = lvwAscending Then
 lvwFabricant.SortOrder = lvwDescending
 Else
 lvwFabricant.SortOrder = lvwAscending
 End If
 
 lvwFabricant.SortKey = ColumnHeader.Index - 1

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "lvwFabricant_ColumnClick", Err, Err.number, Err.Description
End Sub

Private Sub lvwDescription_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

 On Error GoTo Oups

 lvwDescription.Sorted = True

 If lvwDescription.SortOrder = lvwAscending Then
 lvwDescription.SortOrder = lvwDescending
 Else
 lvwDescription.SortOrder = lvwAscending
 End If
 
 lvwDescription.SortKey = ColumnHeader.Index - 1

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "lvwDescription_ColumnClick", Err, Err.number, Err.Description
End Sub

Private Sub lvwFabricant_KeyDown(KeyCode As Integer, Shift As Integer)
 
 On Error GoTo Oups

 If KeyCode = vbKeyReturn Then
 Call lvwFabricant_DblClick
 End If

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "lvwFabricant_KeyDown", Err, Err.number, Err.Description
End Sub

Private Sub lvwPieces_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

 On Error GoTo Oups

 If lvwPieces.SortOrder = lvwAscending Then
 lvwPieces.SortOrder = lvwDescending
 Else
 lvwPieces.SortOrder = lvwAscending
 End If

 lvwPieces.SortKey = ColumnHeader.Index - 1

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "lvwPieces_ColumnClick", Err, Err.number, Err.Description
End Sub

Private Sub lvwPieces_KeyDown(KeyCode As Integer, Shift As Integer)
 
 On Error GoTo Oups

 If KeyCode = vbKeyReturn Then
 Call lvwPieces_DblClick
 End If

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "lvwPieces_KeyDown", Err, Err.number, Err.Description
End Sub



Private Sub lvwRechercheJob_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

 On Error GoTo Oups

 If lvwRechercheJob.SortOrder = lvwAscending Then
 lvwRechercheJob.SortOrder = lvwDescending
 Else
 lvwRechercheJob.SortOrder = lvwAscending
 End If

 lvwRechercheJob.SortKey = ColumnHeader.Index - 1

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "lvwRechercheJob_ColumnClick", Err, Err.number, Err.Description
End Sub

Private Sub lvwRechercheAchat_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

 On Error GoTo Oups

 If lvwRechercheAchat.SortOrder = lvwAscending Then
 lvwRechercheAchat.SortOrder = lvwDescending
 Else
 lvwRechercheAchat.SortOrder = lvwAscending
 End If

 lvwRechercheAchat.SortKey = ColumnHeader.Index - 1

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "lvwRechercheAchat_ColumnClick", Err, Err.number, Err.Description
End Sub

Private Sub lvwFabricant_DblClick()

 On Error GoTo Oups

 Dim itmFabricant As ListItem
 Dim iCompteur As Integer
 
 Screen.MousePointer = vbHourglass
 
 Set itmFabricant = lvwFabricant.SelectedItem

 m_sSelectCategorie = Trim$(itmFabricant.Tag)
 m_sSelectFabricant = Trim$(itmFabricant.Text)
 m_sSelectNoItem = Trim$(itmFabricant.SubItems(I_COL_MAN_PIECE))
 
 Call RemplirComboCategorie
 
 For iCompteur = 0 To cmbCategorie.ListCount - 1
 If cmbCategorie.LIST(iCompteur) = Trim$(itmFabricant.Tag) Then
  cmbCategorie.ListIndex = iCompteur
 
  Exit For
  End If
  Next
 
  For iCompteur = 0 To cmbNoItem.ListCount - 1
  If cmbNoItem.LIST(iCompteur) = Trim$(itmFabricant.SubItems(I_COL_MAN_PIECE)) Then
  cmbNoItem.ListIndex = iCompteur
 
  Exit For
End If
Next
 
lvwFabricant.Visible = False

Screen.MousePointer = vbDefault

Exit Sub

Oups:

wOups "frmCatalogueElec", "lvwFabricant_DblClick", Err, Err.number, Err.Description
End Sub

Private Sub lvwDescription_DblClick()

 On Error GoTo Oups

 Dim itmDescription As ListItem
 Dim iCompteur As Integer

 If lvwDescription.ListItems.count > 0 Then
 Screen.MousePointer = vbHourglass

 Set itmDescription = lvwDescription.SelectedItem

 m_sSelectCategorie = itmDescription.Tag
 m_sSelectFabricant = Trim$(itmDescription.SubItems(I_COL_DES_FABRICANT))
 m_sSelectNoItem = Trim$(itmDescription.SubItems(I_COL_DES_PIECE))

 If m_eMode = MODE_INACTIF Then
 Call RemplirComboCategorie
  Else
  txtDescriptionFR.Text = itmDescription.Text
  txtDescriptionEN.Text = itmDescription.SubItems(I_COL_DES_DESCR_EN)
  End If

  lvwDescription.Visible = False

  Screen.MousePointer = vbDefault
  End If

  Exit Sub

Oups:

10 wOups "frmCatalogueElec", "lvwDescription_DblClick", Err, Err.number, Err.Description
End Sub

Private Sub lvwFabricant_LostFocus()

 On Error GoTo Oups

 If lvwFabricant.Visible = True Then
 lvwFabricant.Visible = False
 End If

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "lvwFabricant_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub lvwPieces_DblClick()

 On Error GoTo Oups

 Dim itmPieces As ListItem
 Dim iCompteur As Integer
 
 Screen.MousePointer = vbHourglass
 
 Set itmPieces = lvwPieces.SelectedItem
 
 m_sSelectCategorie = Trim$(itmPieces.Tag)
 m_sSelectFabricant = Trim$(itmPieces.SubItems(I_COL_PIECE_FABRICANT))
 m_sSelectNoItem = Trim$(itmPieces.Text)
 
 Call RemplirComboCategorie
 
 lvwPieces.Visible = False

 Screen.MousePointer = vbDefault

  Exit Sub

Oups:

  wOups "frmCatalogueElec", "lvwPieces_DblClick", Err, Err.number, Err.Description
End Sub

Private Sub CmdSupp_Click()

 On Error GoTo Oups

 Dim sDescription As String
 Dim iCompteur As Integer

 If cmbNoItem.ListCount > 0 Then
 If MsgBox("Voulez-vous vraiment effacer la pièce " & txtNoItem.Text & "?", vbYesNo) = vbYes Then
 If chkInventaire.Value = vbChecked Then
 Call g_connData.Execute("DELETE * FROM GrbInventaireElec WHERE NoItem = '" & Replace(cmbNoItem.Text, "'", "''") & "'")
 End If

 'Efface l'enregistrement de catalogue
 Call g_connData.Execute("DELETE * FROM GrbCatalogueElec WHERE PIECE = '" & Replace(cmbNoItem.Text, "'", "''") & "'")
 
 'Efface l'enr de la table piece frs
 Call g_connData.Execute("DELETE * FROM GrbPiecesFRS WHERE PIECE = '" & Replace(cmbNoItem.Text, "'", "''") & "'")
 
 m_bRempliManuel = True
 
  m_sSelectNoItem = vbNullString
 
  If cmbNoItem.ListCount > 1 Then
  m_sSelectFabricant = cmbFabricant.Text
  Else
  m_sSelectFabricant = vbNullString
  End If
 
  Call RemplirComboFabricant
 
  If cmbFabricant.ListCount = 0 Then
 Call cmbNoItem.Clear
 
 Call cmbCategorie.RemoveItem(cmbCategorie.ListIndex)
 
 If cmbCategorie.ListCount > 0 Then
 cmbCategorie.ListIndex = 0
 Else
 Call ViderChamps_frs
 
 Call lvwfournisseur.ListItems.Clear
 
 Call ViderChamps_piece
 End If
 End If

 sDescription = txtDescriptionFR.Text

 For iCompteur = 0 To cmbDescriptionFR.ListCount - 1
 If cmbDescriptionFR.LIST(iCompteur) = sDescription Then
 m_bBloqueDescription = True

 cmbDescriptionFR.ListIndex = iCompteur

 m_bBloqueDescription = False

 Exit For
 End If
 Next
1  End If
 End If

 Exit Sub

Oups:

wOups "frmCatalogueElec", "CmdSupp_Click", Err, Err.number, Err.Description
End Sub

Private Sub AfficherItem()

 On Error GoTo Oups

 'Affichage de l'enregistrement
 Dim rstItem As ADODB.Recordset
 Dim rstInventaire As ADODB.Recordset
 Dim iCompteur As Integer
 
 'Il faut mettre le frame enabled pour vérifier si les CheckBox à l'intérieur
 'sont enabled
 Call ViderChamps_piece

 Set rstItem = New ADODB.Recordset

 Call rstItem.Open("SELECT * FROM GrbCatalogueElec WHERE PIECE = '" & Replace(m_sNoItem, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Si il y a un enregistrement
 If Not rstItem.EOF Then
 'PIECE_GRB
 If Not IsNull(rstItem.Fields("PIECE_GRB")) Then
 txtNoItemGRB.Text = Trim(rstItem.Fields("PIECE_GRB"))
 Else
  txtNoItemGRB.Text = vbNullString
  End If

 'DESCR_EN
  If Not IsNull(rstItem.Fields("DESC_EN")) Then
  txtDescriptionEN.Text = Trim(rstItem.Fields("DESC_EN"))
  Else
  txtDescriptionEN.Text = vbNullString
74 End If

 'FABRICANT
  If Not IsNull(rstItem.Fields("FABRICANT").Value) Then
  txtFabricant.Text = Trim(rstItem.Fields("FABRICANT"))
  Else
  txtFabricant.Text = vbNullString
84 End If

 'DESCR_FR
  If Not IsNull(rstItem.Fields("DESC_FR")) Then
 For iCompteur = 0 To cmbDescriptionFR.ListCount - 1
 If cmbDescriptionFR.LIST(iCompteur) = Trim(rstItem.Fields("DESC_FR")) Then
 cmbDescriptionFR.ListIndex = iCompteur
 
 Exit For
 End If
 Next
 Else
 If cmbDescriptionFR.ListIndex = -1 Then
 Call cmbDescriptionFR_Click
 Else
 cmbDescriptionFR.ListIndex = -1
 End If
End If
 
 'TEMPS
 If Not IsNull(rstItem.Fields("TEMPS")) Then
 txtTemps.Text = Trim(rstItem.Fields("TEMPS"))
 Else
 txtTemps.Text = vbNullString
 End If
 
 'COMMENT
 If Not IsNull(rstItem.Fields("COMMENTAIRE")) Then
1  txtComment.Text = Trim(rstItem.Fields("COMMENTAIRE"))
 Else
 txtComment.Text = vbNullString
 End If
 
 Call RemplirListViewFournisseur
 
Else
 Call MsgBox("Impossible de trouver la pièce!", vbOKOnly, "Erreur")
End If
 
Call rstItem.Close
Set rstItem = Nothing

Set rstInventaire = New ADODB.Recordset

Call rstInventaire.Open("SELECT * FROM GrbInventaireElec WHERE NoItem = '" & Replace(txtNoItem.Text, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)

If Not rstInventaire.EOF Then
chkInventaire.Value = vbChecked

 chkBoite.Value = Abs(CInt(rstInventaire.Fields("CommandeParBoite")))

If chkBoite.Value = vbChecked Then
 txtQuantitéBoite.Text = rstInventaire.Fields("QteBoite")
End If

 For iCompteur = 0 To cmbLocalisation.ListCount - 1
 If cmbLocalisation.LIST(iCompteur) = rstInventaire.Fields("Localisation") Then
 cmbLocalisation.ListIndex = iCompteur

 Exit For
End If
 Next

 txtQuantiteStock.Text = rstInventaire.Fields("QuantitéStock")
 chkMinimum.Value = Abs(CInt(rstInventaire.Fields("Minimum")))
 txtQuantiteMinimum.Text = rstInventaire.Fields("QuantitéMinimum")
 txtQuantiteCommande.Text = rstInventaire.Fields("Commande")
End If

Call rstInventaire.Close
Set rstInventaire = Nothing

Exit Sub

Oups:

wOups "frmCatalogueElec", "AfficherItem", Err, Err.number, Err.Description
End Sub

Private Sub AfficherFRS()

 On Error GoTo Oups

 'Affichage de l'enregistrement
 Dim rstItemFRS As ADODB.Recordset
 Dim iCompteur As Integer

 Set rstItemFRS = New ADODB.Recordset

 Call rstItemFRS.Open("SELECT * FROM GrbPiecesFRS WHERE NoEnreg = " & lvwfournisseur.SelectedItem.Tag, g_connData, adOpenDynamic, adLockOptimistic)
 
 'Si le champs est Enabled.. c'est parce que le champs existe dans la table
 
 'DISTRIB
 For iCompteur = 0 To cmbfrs.ListCount - 1
 If cmbfrs.LIST(iCompteur) = lvwfournisseur.SelectedItem.Text Then
 cmbfrs.ListIndex = iCompteur

 Exit For
 End If
 Next
 
 'PERS_RESS
  If Not IsNull(rstItemFRS.Fields("PERS_RESS")) Then
  For iCompteur = 0 To cmbPersRess.ListCount - 1
  If cmbPersRess.ItemData(iCompteur) = rstItemFRS.Fields("PERS_RESS") Then
  cmbPersRess.ListIndex = iCompteur
 
  Exit For
  End If
  Next
  Else
cmbPersRess.ListIndex = -1
End If

 'PRIX_LIST
If Not IsNull(rstItemFRS.Fields("PRIX_LIST")) Then
 txtPrixList.Text = rstItemFRS.Fields("PRIX_LIST")
Else
 txtPrixList.Text = vbNullString
End If
 
 'ESCOMPTE
If Not IsNull(rstItemFRS.Fields("ESCOMPTE")) Then
 mskEscompte.Text = rstItemFRS.Fields("ESCOMPTE")
Else
 mskEscompte.Text = vbNullString
End If

 'PRIX_NET
1  If Not IsNull(rstItemFRS.Fields("PRIX_NET")) Then
 txtPrixNet.Text = rstItemFRS.Fields("PRIX_NET")
 Else
 txtPrixNet.Text = vbNullString
 End If
 
 'PRIX_SP
If Not IsNull(rstItemFRS.Fields("PRIX_SP")) Then
 txtPrixSpecial.Text = rstItemFRS.Fields("PRIX_SP")
1  Else
 txtPrixSpecial.Text = vbNullString
 End If
 
 
 'VALIDE
If Not IsNull(rstItemFRS.Fields("VALIDE")) Then
 mskValide.Text = rstItemFRS.Fields("VALIDE")
Else
 mskValide.Text = vbNullString
End If
 
 'QUOTER
If rstItemFRS.Fields("quoter") = True Then
 chkquoter.Value = vbChecked
Else
 chkquoter.Value = vbUnchecked
End If
 
 'Devise monétaire
2  If rstItemFRS.Fields("DeviseMonétaire") = "CAN" Then
 optCAN.Value = True
2  Else
 If rstItemFRS.Fields("DeviseMonétaire") = "USA" Then
 optUSA.Value = True
 Else
 optSpain.Value = True
 End If
30 End If
 
 'Affiche Drapeau
Call AfficherDrapeau
 
Call rstItemFRS.Close
Set rstItemFRS = Nothing

Exit Sub

Oups:

wOups "frmCatalogueElec", "AfficherFRS", Err, Err.number, Err.Description
End Sub

Private Sub cmbNoItem_Click()

 On Error GoTo Oups

 'Affichage de l'enregistrement
 Screen.MousePointer = vbHourglass
 
 'Il faut mettre le nom de l'élément sélectionné dans le textbox pour ensuite
 'l'utiliser pour les requêtes SQL
 txtNoItem.Text = cmbNoItem.Text
 
 m_sNoItem = txtNoItem.Text
 
 m_bBloqueDescription = True
 
 Call AfficherItem
 
 m_bBloqueDescription = False
 
 'Remplir combo frs
 Call RemplirComboFRS
 
 Screen.MousePointer = vbDefault

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "cmbNoItem_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmbFabricant_Click()

 On Error GoTo Oups

 'quand un manufact est selectionné on remplir le combo des NumItem
 Screen.MousePointer = vbHourglass
 
 txtFabricant.Text = cmbFabricant.Text
 
 Call RemplirComboDescription
 
 m_bBloqueDescription = True
 
 If m_bRempliManuel = True Then
 
 Call RemplirComboNoItem
 
 m_bRempliManuel = False
 Else
 
 Call RemplirComboNoItem
 End If

  m_bBloqueDescription = False
 
  Screen.MousePointer = vbDefault
 If sChoisirTous = ")" Then
 Call RemplirComboCategorie
 End If
  Exit Sub

Oups:

  wOups "frmCatalogueElec", "cmbFabricant_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdSuppFrs_Click()

 On Error GoTo Oups

 If lvwfournisseur.ListItems.count > 0 Then
 Call SupprimerFournisseur
 End If

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "cmdSuppFrs_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups

 Call Unload(frmChoixCatalogue)

 Call ActiverBoutonsGroupe

 m_bBloqueDescription = True

 Set m_collPieceDescFR = New Collection
 
 'Barrer les champs
 Call BarrerChamps_piece(True)
 
 'Activer ou désactiver certains controles
 Call MontrerControles(MODE_INACTIF)
 
 Call RemplirComboLocalisation

 'Rempli le combo des pièces disponibles
 Call RemplirComboCategorie

 m_bBloqueDescription = False

 Exit Sub

Oups:

  wOups "frmCatalogueElec", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub ActiverBoutonsGroupe()

 On Error GoTo Oups

 CmdAdd.Enabled = g_bModificationCatalogueElec
 CmdSupp.Enabled = g_bModificationCatalogueElec
 CmdModif.Enabled = g_bModificationCatalogueElec
 cmdAddFrs.Enabled = g_bModificationCatalogueElec
 cmdSuppFrs.Enabled = g_bModificationCatalogueElec
 cmdModifFrs.Enabled = g_bModificationCatalogueElec
 cmdDemande.Enabled = g_bModificationCatalogueElec
 
 Exit Sub

Oups:

 wOups "frmCatalogueElec", "ActiverBoutonsGroupe", Err, Err.number, Err.Description
End Sub

Public Sub RemplirComboFabricant()

 On Error GoTo Oups

 'Rempli le combo des fabricants
 Dim rstFabricant As ADODB.Recordset
 Dim sCategorie As String
 Dim iCompteur As Integer
 
 sCategorie = Replace(cmbCategorie.Text, "'", "''")
 
 Set rstFabricant = New ADODB.Recordset
 
 Call rstFabricant.Open("SELECT DISTINCT FABRICANT FROM GrbCatalogueElec WHERE CATEGORIE = '" & sCategorie & "' ORDER BY FABRICANT", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Il faut vider le combo avant de le remplir
 Call cmbFabricant.Clear
4 sChoisirTous = ""
4 'on ajoute la possibilité de choisir tout les fabricants
44 Call cmbFabricant.AddItem("-- CHOISIR TOUS --")
 If Not rstFabricant.EOF Then
 rstFabricant.MoveFirst
 End If
 'Tant que ce n'est pas la fin des enregistrements
 
4  Do While Not rstFabricant.EOF
 'Si l'élément n'est pas null
 
 If Not IsNull(rstFabricant.Fields("Fabricant")) Then
 
 'on l'ajoute
 Call cmbFabricant.AddItem(Trim(rstFabricant.Fields("FABRICANT")))
 If sChoisirTous = "" Then
 sChoisirTous = " AND (FABRICANT = '" & Trim(rstFabricant.Fields("FABRICANT")) & "'"
 Else
5  sChoisirTous = sChoisirTous + " OR FABRICANT = '" & Trim(rstFabricant.Fields("FABRICANT")) & "'"
 End If
  End If
 
  Call rstFabricant.MoveNext
 
  Loop
 
 sChoisirTous = sChoisirTous + ")"
 
  Call rstFabricant.Close
  Set rstFabricant = Nothing
 
 'Si le combo n'est pas vide, on sélectionne le premier élément
  If cmbFabricant.ListCount > 0 Then
 
  If m_sSelectFabricant <> vbNullString Then
  For iCompteur = 0 To cmbFabricant.ListCount - 1
 
 If UCase(cmbFabricant.LIST(iCompteur)) = UCase(m_sSelectFabricant) Then
 cmbFabricant.ListIndex = iCompteur
 
 m_sSelectFabricant = ""
 
 Exit For
 End If
 Next
 
 Else
 
 cmbFabricant.ListIndex = 0
 
 End If
Else
 
 Call cmbNoItem.Clear
 Call cmbDescriptionFR.Clear
1  End If
 
Exit Sub

Oups:

 wOups "frmCatalogueElec", "RemplirComboFabricant", Err, Err.number, Err.Description
End Sub

Public Sub AfficherForm(ByVal sCategorie As String, ByVal sNomFab As String, ByVal sNoItem As String)

 On Error GoTo Oups


 Dim iCompteur As Integer
 'Ouverture de la fenêtre
 
 'Barrer les champs
 Call BarrerChamps_piece(True)
 
 'Activer ou désactiver certains controles
 Call MontrerControles(MODE_INACTIF)
 
 'Remplir le combo des pièces disponibles
 Call RemplirComboCategorie
 
 If sCategorie <> "" Then
 For iCompteur = 0 To cmbCategorie.ListCount - 1
 If cmbCategorie.LIST(iCompteur) = sCategorie Then
 cmbCategorie.ListIndex = iCompteur

 Exit For
 End If
  Next
  End If
 
  If sNomFab <> "" Then
  For iCompteur = 0 To cmbFabricant.ListCount - 1
  If cmbFabricant.LIST(iCompteur) = sNomFab Then
  cmbFabricant.ListIndex = iCompteur

  Exit For
  End If
Next
End If

If sNoItem <> "" Then
 For iCompteur = 0 To cmbNoItem.ListCount - 1
 If cmbNoItem.LIST(iCompteur) = sNoItem Then
 cmbNoItem.ListIndex = iCompteur

 Exit For
 End If
 Next
End If
 
Call Me.Show

Exit Sub

Oups:

1  wOups "frmCatalogueElec", "AfficherForm", Err, Err.number, Err.Description
End Sub

Public Sub RemplirComboNoItem()

 On Error GoTo Oups

 'Rempli le combo de numéros d'item
 Dim rstNoItem As ADODB.Recordset
 Dim sCategorie As String
 Dim iCompteur As Integer
 Dim sFabricant As String
 
 sCategorie = Replace(cmbCategorie.Text, "'", "''")
 sFabricant = Replace(cmbFabricant.Text, "'", "''")
 
 Set rstNoItem = New ADODB.Recordset
 If cmbCategorie.Text = "DIVERS" Or sChoisirTous = ")" Then
 Call rstNoItem.Open("SELECT PIECE FROM GrbCatalogueElec WHERE CATEGORIE = '" & sCategorie & "' ORDER BY TRIM(PIECE)", g_connData, adOpenDynamic, adLockOptimistic)
4 Else
 If sFabricant = "-- CHOISIR TOUS --" Then
4 Call rstNoItem.Open("SELECT PIECE FROM GrbCatalogueElec WHERE CATEGORIE = '" & sCategorie & "'" & sChoisirTous & " ORDER BY TRIM(PIECE)", g_connData, adOpenDynamic, adLockOptimistic)
4 Else
44 Call rstNoItem.Open("SELECT PIECE FROM GrbCatalogueElec WHERE CATEGORIE = '" & sCategorie & "' AND FABRICANT = '" & sFabricant & "' ORDER BY TRIM(PIECE)", g_connData, adOpenDynamic, adLockOptimistic)
 End If
 End If
 'Il faut vider le combo avant de le remplir
 Call cmbNoItem.Clear

 'Tant que c'est n'est pas la fin des enregistrements
 Do While Not rstNoItem.EOF
 'Si le champs n'est pas vide
  If Not IsNull(rstNoItem.Fields("PIECE")) Then
 'On l'ajoute
  Call cmbNoItem.AddItem(Trim(rstNoItem.Fields("PIECE")))
  End If
 
  Call rstNoItem.MoveNext
  Loop
 
  Call rstNoItem.Close
  Set rstNoItem = Nothing
 
 'Si le combo n'est pas vide, on sélectionne le premier élément
  If cmbNoItem.ListCount > 0 Then
If m_sSelectNoItem <> vbNullString Then
For iCompteur = 0 To cmbNoItem.ListCount - 1
 If cmbNoItem.LIST(iCompteur) = m_sSelectNoItem Then
 cmbNoItem.ListIndex = iCompteur
 
 m_sSelectNoItem = ""
 
 Exit For
 End If
 Next
 Else
 
 cmbNoItem.ListIndex = 0
 End If
End If

1  Exit Sub

Oups:

wOups "frmCatalogueElec", "RemplirComboNoItem", Err, Err.number, Err.Description
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
  sPrixCalcul = Replace(rstPieceFRS.Fields("PRIX_NET"), ".", ",")
  Else
 If rstPieceFRS.Fields("PRIX_SP") <> vbNullString Then
 sPrixCalcul = Replace(rstPieceFRS.Fields("PRIX_SP"), ".", ",")
 End If
 End If
 
 If rstPieceFRS.Fields("DeviseMonétaire") = "USA" Then
 rstPieceFRS.Fields("PrixReel") = Conversion(CStr(Round(CDbl(sPrixCalcul) / CDbl(sTauxUSA), 4)), MODE_DECIMAL, 4)
 Else
 If rstPieceFRS.Fields("DeviseMonétaire") = "SPA" Then
 rstPieceFRS.Fields("PrixReel") = Conversion(CStr(Round(CDbl(sPrixCalcul) / CDbl(sTauxSPA), 4)), MODE_DECIMAL, 4)
 Else
 rstPieceFRS.Fields("PrixReel") = Conversion(sPrixCalcul, MODE_DECIMAL, 4)
 End If
End If
 
 Call rstPieceFRS.Update
 
 Call rstPieceFRS.MoveNext
Loop
 
 Call rstPieceFRS.Close
Set rstPieceFRS = Nothing

 Exit Sub

Oups:

1  wOups "frmCatalogueElec", "CalculerPrixReel", Err, Err.number, Err.Description
End Sub

Public Sub RemplirListViewFournisseur()

 On Error GoTo Oups

 ''''''''''''''''''''''''''''''
 ' remplis lister fournisseur '
 ''''''''''''''''''''''''''''''
 Dim rstPieceFRS As ADODB.Recordset
 Dim rstContact As ADODB.Recordset
 Dim iCompteur As Integer
 Dim itmFRS As ListItem
 Dim lCouleur As Long
 
 'vide le lister
 Call lvwfournisseur.ListItems.Clear
 
 Call CalculerPrixReel(txtNoItem.Text)
 
 Set rstPieceFRS = New ADODB.Recordset
 
 Call rstPieceFRS.Open("SELECT GrbPiecesFRS.*, GrbFournisseur.NomFournisseur FROM GrbPiecesFRS INNER JOIN GrbFournisseur ON GrbPiecesFRS.IDFRS = GrbFournisseur.IDFRS WHERE GrbPiecesFRS.PIECE = '" & Replace(txtNoItem.Text, "'", "''") & "' AND Type = 'E' ORDER BY PrixReel", g_connData, adOpenDynamic, adLockOptimistic)
 
 Set rstContact = New ADODB.Recordset
 
 'tant il y a des fournisseur de la piece , ajoute dans lister
 
  Do While Not rstPieceFRS.EOF
  If rstPieceFRS.Fields("DeviseMonétaire") = "CAN" Then
  lCouleur = COLOR_ROUGE
 
  Else
  lCouleur = COLOR_BLEU
 
  End If
 
  Set itmFRS = lvwfournisseur.ListItems.Add
 
  itmFRS.Text = rstPieceFRS.Fields("NomFournisseur")
itmFRS.ForeColor = lCouleur
 
1 itmFRS.Tag = rstPieceFRS.Fields("NoEnreg")
 
 If Not IsNull(rstPieceFRS.Fields("PERS_RESS")) Then
 If Trim(rstPieceFRS.Fields("PERS_RESS")) <> vbNullString Then
 Call rstContact.Open("SELECT NomContact FROM GrbContact WHERE IDContact = " & rstPieceFRS.Fields("PERS_RESS"), g_connData, adOpenDynamic, adLockOptimistic)
 
 If Not rstContact.EOF Then
 itmFRS.SubItems(I_COL_FRS_PERS_RESS) = rstContact.Fields("NomContact")
 itmFRS.ListSubItems(I_COL_FRS_PERS_RESS).ForeColor = lCouleur
 End If
 
 Call rstContact.Close
 End If
 End If
 
 If Not IsNull(rstPieceFRS.Fields("Date")) Then
 
 itmFRS.SubItems(I_COL_FRS_DATE) = rstPieceFRS.Fields("Date")
Else
 
 itmFRS.SubItems(I_COL_FRS_DATE) = vbNullString
 End If
 
 itmFRS.ListSubItems(I_COL_FRS_DATE).ForeColor = lCouleur
 
 If Not IsNull(rstPieceFRS.Fields("ENTRER_PAR")) Then
185
 itmFRS.SubItems(I_COL_FRS_ENTRER_PAR) = rstPieceFRS.Fields("ENTRER_PAR")
 Else
 
1  itmFRS.SubItems(I_COL_FRS_ENTRER_PAR) = vbNullString
 End If
 
 itmFRS.ListSubItems(I_COL_FRS_ENTRER_PAR).ForeColor = lCouleur
 
 If Not IsNull(rstPieceFRS.Fields("Valide")) Then
 
 itmFRS.SubItems(I_COL_FRS_VALIDE) = rstPieceFRS.Fields("Valide")
 Else
 
 itmFRS.SubItems(I_COL_FRS_VALIDE) = vbNullString
 End If
 
 itmFRS.ListSubItems(I_COL_FRS_VALIDE).ForeColor = lCouleur
 
 If rstPieceFRS.Fields("PRIX_LIST") <> vbNullString Then
 itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = Conversion(Round(Replace(rstPieceFRS.Fields("PRIX_LIST"), ".", ","), 4), MODE_ARGENT, 4)

 itmFRS.ListSubItems(I_COL_FRS_PRIX_LIST).ForeColor = lCouleur
 End If
 
If Not IsNull(rstPieceFRS.Fields("ESCOMPTE")) Then
 If Trim(rstPieceFRS.Fields("ESCOMPTE")) <> vbNullString Then
 'Enlève les "_", met un format pourcentage et remplace les "." par des ","
 itmFRS.SubItems(I_COL_FRS_ESCOMPTE) = Conversion(Replace(Replace(rstPieceFRS.Fields("ESCOMPTE"), "_", vbNullString), ".", ","), MODE_POURCENT)
 Else
 itmFRS.SubItems(I_COL_FRS_ESCOMPTE) = vbNullString
 End If
Else
 itmFRS.SubItems(I_COL_FRS_ESCOMPTE) = vbNullString
End If
 
3 itmFRS.ListSubItems(I_COL_FRS_ESCOMPTE).ForeColor = lCouleur
 
 If rstPieceFRS.Fields("PRIX_NET") <> vbNullString Then
 itmFRS.SubItems(I_COL_FRS_PRIX_NET) = Conversion(Round(Replace(rstPieceFRS.Fields("PRIX_NET"), ".", ","), 4), MODE_ARGENT, 4)
 Else
 itmFRS.SubItems(I_COL_FRS_PRIX_NET) = vbNullString
 End If
 
 itmFRS.ListSubItems(I_COL_FRS_PRIX_NET).ForeColor = lCouleur
 
 If rstPieceFRS.Fields("PRIX_SP") <> vbNullString Then
 itmFRS.SubItems(I_COL_FRS_PRIX_SP) = Conversion(Round(rstPieceFRS.Fields("PRIX_SP"), 4), MODE_ARGENT, 4)
 Else
 itmFRS.SubItems(I_COL_FRS_PRIX_SP) = vbNullString
End If
 
 itmFRS.ListSubItems(I_COL_FRS_PRIX_SP).ForeColor = lCouleur
 
If rstPieceFRS.Fields("QUOTER") = True Then
 itmFRS.SubItems(I_COL_FRS_QUOTER) = "Oui"
Else
 itmFRS.SubItems(I_COL_FRS_QUOTER) = "Non"
 End If
 
 itmFRS.ListSubItems(I_COL_FRS_QUOTER).ForeColor = lCouleur
 
Call rstPieceFRS.MoveNext
Loop
 
 'Ferme la table
4 Call rstPieceFRS.Close
4 Set rstPieceFRS = Nothing

4 Set rstContact = Nothing

4 Exit Sub

Oups:

4 wOups "frmCatalogueElec", "RemplirListViewFournisseur", Err, Err.number, Err.Description
End Sub

Private Sub lvwDescription_LostFocus()

 On Error GoTo Oups

 If lvwDescription.Visible = True Then
 lvwDescription.Visible = False
 End If

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "lvwDescription_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub lvwRechercheJob_LostFocus()

 On Error GoTo Oups

 If lvwRechercheJob.Visible = True Then
 lvwRechercheJob.Visible = False
 End If

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "lvwRechercheJob_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub lvwRechercheAchat_LostFocus()

 On Error GoTo Oups

 If lvwRechercheAchat.Visible = True Then
 lvwRechercheAchat.Visible = False
 End If

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "lvwRechercheAchat_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub lvwFournisseur_DblClick()

 On Error GoTo Oups

 'modifie un fournisseur pour la piece
 If lvwfournisseur.ListItems.count > 0 Then
 Call ModifierFournisseur
 End If

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "lvwFournisseur_DblClick", Err, Err.number, Err.Description
End Sub

Private Sub lvwfournisseur_KeyDown(KeyCode As Integer, Shift As Integer)

 On Error GoTo Oups

 If lvwfournisseur.ListItems.count > 0 Then
 If KeyCode = vbKeyDelete Then
 Call SupprimerFournisseur
 End If
 End If

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "lvwfournisseur_KeyDown", Err, Err.number, Err.Description
End Sub

Private Sub ModifierFournisseur()

 On Error GoTo Oups

 Call BarrerChamps_piece(True)
 
 'affiche pour entre des valeurs
 Call MontrerControles(MODE_AJOUT_MODIF_FRS)

 m_bAjout = False

 'affiche les données frs selectionné
 Call AfficherFRS

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "ModifierFournisseur", Err, Err.number, Err.Description
End Sub

Private Sub SupprimerFournisseur()

 On Error GoTo Oups

 If MsgBox("Voulez-vous vraiment effacer le fournisseur " & lvwfournisseur.SelectedItem.Text & "?", vbYesNo) = vbYes Then
 'fonction qui supprime l'enregistrer courant
 Call g_connData.Execute("DELETE * FROM GrbPiecesFRS WHERE NoEnreg = " & lvwfournisseur.SelectedItem.Tag)
 
 'remplir le lister des fournisseurs
 Call RemplirListViewFournisseur
 End If

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "SupprimerFournisseur", Err, Err.number, Err.Description
End Sub

Private Sub lvwPieces_LostFocus()

 On Error GoTo Oups

 If lvwPieces.Visible = True Then
 lvwPieces.Visible = False
 End If

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "lvwPieces_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskEscompte_GotFocus()

 On Error GoTo Oups

 'Quand le maskEdit prend le focus, on set le masque
 If mskEscompte.Enabled = True Then
 mskEscompte.mask = "0,####"
 End If

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "mskEscompte_GotFocus", Err, Err.number, Err.Description
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

 wOups "frmCatalogueElec", "mskEscompte_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub optCAN_Click()

 On Error GoTo Oups

 'dependant la devise, affiche le drapeau
 Call AfficherDrapeau

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "optCAN_Click", Err, Err.number, Err.Description
End Sub
 
Private Sub AfficherDrapeau()

 On Error GoTo Oups

 'dependant la devise, affiche le drapeau
 If optCAN.Value = True Then
 imgCanada.Visible = True
 imgEU.Visible = False
 imgSpain.Visible = False


 lblDevise1.Visible = False
 txtTauxChange.Visible = False
 lblDevise2.Visible = False
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

  Call AfficherTauxChange
10 End If

Exit Sub

Oups:

wOups "frmCatalogueElec", "AfficherDrapeau", Err, Err.number, Err.Description
End Sub

Private Sub AfficherTauxChange()

 On Error GoTo Oups

 Dim rstConfig As ADODB.Recordset

 Set rstConfig = New ADODB.Recordset

 Call rstConfig.Open("SELECT TauxAmericain, TauxEspagnol FROM GrbConfig", g_connData, adOpenDynamic, adLockOptimistic)

 If optUSA.Value = True Then
 lblDevise2.Caption = "$ USA"
 txtTauxChange.Text = rstConfig.Fields("TauxAmericain")
 Else
 lblDevise2.Caption = "$ SPA"
 txtTauxChange.Text = rstConfig.Fields("TauxEspagnol")
 End If

  lblDevise1.Visible = True
  txtTauxChange.Visible = True
  lblDevise2.Visible = True

  Call rstConfig.Close
  Set rstConfig = Nothing

  Exit Sub

Oups:

  wOups "frmCatalogueElec", "AfficherTauxChange", Err, Err.number, Err.Description
End Sub

Private Sub optSpain_Click()

 On Error GoTo Oups

 'dependant la devise, affiche le drapeau
 Call AfficherDrapeau

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "optSpain_Click", Err, Err.number, Err.Description
End Sub

Private Sub optUSA_Click()

 On Error GoTo Oups

 'dependant la devise, affiche le drapeau
 Call AfficherDrapeau

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "optUSA_Click", Err, Err.number, Err.Description
End Sub

Private Sub txtNoItem_Change()

 On Error GoTo Oups

 If m_eMode = MODE_AJOUT_MODIF_ELEC Then
 If Len(txtNoItem.Text) > 1 Then
 txtNoItemGRB.Text = Left$(txtNoItem.Text, 18) & "GRB"
 Else
 txtNoItemGRB.Text = txtNoItem.Text & "GRB"
 End If
 End If

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "txtNoItem_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtPrixList_LostFocus()

 On Error GoTo Oups

 If txtPrixList.Text <> vbNullString Then
 txtPrixList.Text = Replace(txtPrixList, ".", ",")
 
 If Not IsNumeric(txtPrixList.Text) Then
 Call MsgBox("Valeur non numérique!", vbOKOnly, "Erreur")
 txtPrixList.Text = vbNullString
 End If
 End If

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "txtPrixList_LostFocus", Err, Err.number, Err.Description
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

 wOups "frmCatalogueElec", "txtPrixNet_Change", Err, Err.number, Err.Description
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

  wOups "frmCatalogueElec", "txtPrixSpecial_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtPrixNet_GotFocus()

 On Error GoTo Oups

 'Si le prix net prend le focus
 Call CalculerPrixNet

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "txtPrixNet_GotFocus", Err, Err.number, Err.Description
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

wOups "frmCatalogueElec", "CalculerPrixNet", Err, Err.number, Err.Description
End Sub

Private Sub txtPrixNet_LostFocus()

 On Error GoTo Oups

 txtPrixNet.Text = Replace(txtPrixNet.Text, ".", ",")
 
 Exit Sub

Oups:

 wOups "frmCatalogueElec", "txtPrixNet_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskValide_GotFocus()

 On Error GoTo Oups

 'Si la date est sous le format AAAA-MM-JJ
 If Len(mskValide.Text) = 10 Then
 'On la met sous le format AA-MM-JJ
 mskValide.Text = Right$(mskValide.Text, 8)
 End If
 
 'On met le mask
 mskValide.mask = "##-##-##"

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "mskValide_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskValide_LostFocus()

 On Error GoTo Oups
 'On enlève le mask
 mskValide.mask = vbNullString
 
 If mskValide.Text = "__-__-__" Then
 mskValide.Text = vbNullString
 Else
 If Len(mskValide.Text) =   Then
 If IsDate(mskValide.Text) Then
 'On la met sous le format AAAA-MM-JJ
 mskValide.Text = Year(DateSerial(Left$(mskValide.Text, 2), Mid$(mskValide.Text, 4, 2), Right$(mskValide.Text, 2))) & Mid$(mskValide.Text, 3, 8)
 End If
 End If
 End If

  Exit Sub

Oups:

  wOups "frmCatalogueElec", "mskValide_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub cmbCategorie_Click()
 
 On Error GoTo Oups

 'pour sélectionner la bonne catégorie de pieces
 txtCategorie.Text = cmbCategorie.Text
 
 m_bRempliManuel = True

 m_bBloqueDescription = True
 
 Call cmbFabricant.Clear
 
 Call cmbNoItem.Clear
 
 Call ViderChamps_piece
 
 Call RemplirComboFabricant
 
 m_bBloqueDescription = False

 Screen.MousePointer = vbDefault
 
 Exit Sub

Oups:

  wOups "frmCatalogueElec", "cmbCategorie_Click", Err, Err.number, Err.Description
End Sub

Public Sub RemplirComboCategorie()

 On Error GoTo Oups

 'Remplir le combo des tables (Pièces)
 Dim rstCatalogueElec As ADODB.Recordset
 Dim iCompteur As Integer
 
 'Il faut vider le combo avant de le remplir
 Call cmbCategorie.Clear
 
 'Cette méthode crée un recordset contenant les categorie
 'le nom de toutes les tables de la BD
 Set rstCatalogueElec = New ADODB.Recordset
 
 Call rstCatalogueElec.Open("SELECT DISTINCT CATEGORIE FROM GrbCatalogueElec ORDER BY CATEGORIE", g_connData, adOpenDynamic, adLockOptimistic)

 'Tant que ce n'est pas la fin des enregistrements
 Do While Not rstCatalogueElec.EOF
 If Not IsNull(rstCatalogueElec.Fields("CATEGORIE")) Then
 Call cmbCategorie.AddItem(Trim(rstCatalogueElec.Fields("CATEGORIE")))
 End If
 
 Call rstCatalogueElec.MoveNext
  Loop
 
  Call rstCatalogueElec.Close
  Set rstCatalogueElec = Nothing
 
 'Si le combo n'est pas vide, on sélectionne le premier
  If cmbCategorie.ListCount > 0 Then
  If m_sSelectCategorie <> "" Then
  For iCompteur = 0 To cmbCategorie.ListCount - 1
  If cmbCategorie.LIST(iCompteur) = m_sSelectCategorie Then
  cmbCategorie.ListIndex = iCompteur

 m_sSelectCategorie = ""

 Exit For
 End If
 Next
 Else
 cmbCategorie.ListIndex = 0
 End If
End If

Exit Sub

Oups:

wOups "frmCatalogueElec", "RemplirComboCategorie", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboFRS()

 On Error GoTo Oups

 'Remplir le combo des tables (Pièces)
 Dim rstPieceFRS As ADODB.Recordset
 Dim sNomTable As String
 
 'Il faut vider le combo avant de le remplir
 Call cmbfrs.Clear
 
 ' ouvre la table piece frs
 Set rstPieceFRS = New ADODB.Recordset
 
 Call rstPieceFRS.Open("SELECT * FROM GrbFournisseur WHERE Supprimé = False ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Tant que ce n'est pas la fin des enregistrements
 
 Do While Not rstPieceFRS.EOF
 Call cmbfrs.AddItem(rstPieceFRS.Fields("NomFournisseur"))
 cmbfrs.ItemData(cmbfrs.newIndex) = rstPieceFRS.Fields("IDFRS")
 
 Call rstPieceFRS.MoveNext
 Loop
 
  Call rstPieceFRS.Close
  Set rstPieceFRS = Nothing

  Exit Sub

Oups:

  wOups "frmCatalogueElec", "RemplirComboFRS", Err, Err.number, Err.Description
End Sub

Private Sub txtPrixSpecial_LostFocus()

 On Error GoTo Oups

 txtPrixSpecial.Text = Replace(txtPrixSpecial.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "txtPrixSpecial_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboLocalisation()

 On Error GoTo Oups

 'Rempli le combo cmbLocalisation
 Dim rstLocalisation As ADODB.Recordset
 
 Set rstLocalisation = New ADODB.Recordset
 
 Call rstLocalisation.Open("SELECT DISTINCT Localisation FROM GrbInventaireElec", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Il faut vider le combo avant de le remplir
 Call cmbLocalisation.Clear
 
 'Tant que ce n'est pas la fin des enregistrements
 Do While Not rstLocalisation.EOF
 'Si l'enregistrement n'est pas Null
 If Not IsNull(rstLocalisation.Fields("Localisation")) Then
 If Trim(rstLocalisation.Fields("Localisation")) <> "" Then
 'On l'ajoute dans le combo
 Call cmbLocalisation.AddItem(rstLocalisation.Fields("Localisation"))
 End If
 End If
 
  Call rstLocalisation.MoveNext
  Loop
 
  Call rstLocalisation.Close
  Set rstLocalisation = Nothing

  Exit Sub

Oups:

  wOups "frmCatalogueElec", "RemplirComboLocalisation", Err, Err.number, Err.Description
End Sub

Private Sub txtQuantitéBoite_LostFocus()

 On Error GoTo Oups

 txtQuantitéBoite.Text = Replace(txtQuantitéBoite.Text, ".", ",")

 If Not IsNumeric(txtQuantitéBoite.Text) Or txtQuantitéBoite.Text = "0" Then
 txtQuantitéBoite.Text = "1"
 End If

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "txtQuantitéBoite_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtQuantiteCommande_LostFocus()

 On Error GoTo Oups

 txtQuantiteCommande.Text = Replace(txtQuantiteCommande.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "txtQuantiteCommande_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtQuantiteMinimum_LostFocus()

 On Error GoTo Oups

 txtQuantiteMinimum.Text = Replace(txtQuantiteMinimum.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "txtQuantiteMinimum_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub txtQuantiteStock_LostFocus()

 On Error GoTo Oups

 txtQuantiteStock.Text = Replace(txtQuantiteStock.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmCatalogueElec", "txtQuantiteStock_LostFocus", Err, Err.number, Err.Description
End Sub
