VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCatalogueMec 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catalogue Mécanique"
   ClientHeight    =   6615
   ClientLeft      =   1455
   ClientTop       =   1020
   ClientWidth     =   10410
   Icon            =   "FrmCatalogueMec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6615
   ScaleWidth      =   10410
   Begin MSComctlLib.ListView lvwCategorie 
      Height          =   1935
      Left            =   5280
      TabIndex        =   82
      Top             =   120
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   3413
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
         Text            =   "Catégorie"
         Object.Width           =   6703
      EndProperty
   End
   Begin VB.CommandButton cmdRechercheCategorie 
      Height          =   375
      Left            =   9960
      Picture         =   "FrmCatalogueMec.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   120
      Width           =   375
   End
   Begin MSComctlLib.ListView lvwRechercheAchat 
      Height          =   1935
      Left            =   1560
      TabIndex        =   79
      Top             =   960
      Visible         =   0   'False
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   3413
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
      Height          =   1935
      Left            =   1560
      TabIndex        =   15
      Top             =   960
      Visible         =   0   'False
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   3413
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
   Begin VB.CommandButton cmdRechercheJob 
      Caption         =   "Recherche dans jobs / soums"
      Height          =   495
      Left            =   1080
      TabIndex        =   38
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton cmdCopier 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Copier"
      Height          =   495
      Left            =   3240
      TabIndex        =   69
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      Height          =   375
      Left            =   120
      TabIndex        =   37
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdRechercheInventaire 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Inventaire"
      Height          =   495
      Left            =   120
      TabIndex        =   67
      Top             =   6000
      Width           =   1455
   End
   Begin MSComctlLib.ListView lvwFabricant 
      Height          =   1935
      Left            =   1560
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   3413
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
      Height          =   1935
      Left            =   1560
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   3413
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
   Begin VB.CommandButton cmdChangerCategorie 
      Caption         =   "Changer de catégorie"
      Height          =   375
      Left            =   7680
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdDemande 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Demande de prix"
      Height          =   495
      Left            =   1680
      TabIndex        =   68
      Top             =   6000
      Width           =   1455
   End
   Begin MSComctlLib.ListView lvwDescription 
      Height          =   1935
      Left            =   1560
      TabIndex        =   6
      Top             =   960
      Visible         =   0   'False
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   3413
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
         Text            =   "Description anglais"
         Object.Width           =   6138
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Manufacturier"
         Object.Width           =   2090
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "NoItem"
         Object.Width           =   3254
      EndProperty
   End
   Begin VB.CommandButton cmdRechercheDescrFR 
      Height          =   375
      Left            =   9960
      Picture         =   "FrmCatalogueMec.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   900
      Width           =   375
   End
   Begin VB.Frame frafournisseur 
      BackColor       =   &H00000000&
      Caption         =   "Fournisseur"
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   120
      TabIndex        =   41
      Top             =   3480
      Width           =   10215
      Begin VB.TextBox txtTauxChange 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   8640
         TabIndex        =   76
         Top             =   1920
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSComctlLib.ListView lvwfournisseur 
         Height          =   1575
         Left            =   120
         TabIndex        =   42
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
         TabIndex        =   62
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdSuppFrs 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Supprimer"
         Height          =   375
         Left            =   1320
         TabIndex        =   65
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdModifFrs 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Modifier"
         Height          =   375
         Left            =   2520
         TabIndex        =   66
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
         TabIndex        =   61
         Top             =   1440
         Width           =   1335
      End
      Begin VB.OptionButton optUSA 
         BackColor       =   &H00000000&
         Caption         =   "USA"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   8640
         TabIndex        =   56
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton optCAN 
         BackColor       =   &H00000000&
         Caption         =   "CAN"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7920
         TabIndex        =   55
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
         TabIndex        =   48
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtPrixNet 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6360
         TabIndex        =   52
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtPrixSpecial 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6360
         TabIndex        =   54
         Top             =   1440
         Width           =   855
      End
      Begin MSMask.MaskEdBox mskValide 
         Height          =   255
         Left            =   1560
         TabIndex        =   46
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
         TabIndex        =   50
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
         TabIndex        =   63
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdAnnulFrs 
         Caption         =   "A&nnuler"
         Height          =   375
         Left            =   1320
         TabIndex        =   64
         Top             =   1920
         Width           =   1095
      End
      Begin VB.ComboBox cmbfrs 
         Height          =   315
         Left            =   1560
         TabIndex        =   44
         Top             =   360
         Width           =   2775
      End
      Begin VB.ComboBox cmbPersRess 
         Height          =   315
         Left            =   1560
         TabIndex        =   45
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton optSpain 
         BackColor       =   &H00000000&
         Caption         =   "SPA"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   9360
         TabIndex        =   57
         Top             =   360
         Width           =   735
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
         TabIndex        =   78
         Top             =   1920
         Visible         =   0   'False
         Width           =   855
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
         Left            =   9555
         TabIndex        =   77
         Top             =   1920
         Visible         =   0   'False
         Width           =   570
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
         TabIndex        =   60
         Top             =   1080
         Width           =   975
      End
      Begin VB.Image imgCanada 
         Height          =   1065
         Left            =   8160
         Picture         =   "FrmCatalogueMec.frx":0646
         Stretch         =   -1  'True
         Top             =   720
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Image imgEU 
         Height          =   1065
         Left            =   8160
         Picture         =   "FrmCatalogueMec.frx":56628
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
         TabIndex        =   43
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
         TabIndex        =   47
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
         TabIndex        =   49
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
         TabIndex        =   58
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
         TabIndex        =   51
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
         TabIndex        =   53
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
         TabIndex        =   59
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Image imgSpain 
         Height          =   1065
         Left            =   8160
         Picture         =   "FrmCatalogueMec.frx":A339A
         Stretch         =   -1  'True
         Top             =   720
         Visible         =   0   'False
         Width           =   1680
      End
   End
   Begin VB.ComboBox cmbCategorie 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5280
      TabIndex        =   0
      Text            =   "cmbCategorie"
      Top             =   120
      Width           =   4215
   End
   Begin VB.TextBox txtNoItemGRB 
      BackColor       =   &H00FFFFFF&
      DataField       =   "PIECE_GRB"
      DataSource      =   "DatCat1"
      Height          =   285
      Left            =   1560
      TabIndex        =   24
      Top             =   2280
      Width           =   1935
   End
   Begin VB.ComboBox cmbNoItem 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1560
      TabIndex        =   17
      Text            =   "cmbNoItem"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton CmdModif 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Modifier"
      Height          =   495
      Left            =   7560
      TabIndex        =   74
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton CmdFerme 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Fermer"
      Height          =   495
      Left            =   9000
      TabIndex        =   75
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton CmdSupp 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Supprimer"
      Height          =   495
      Left            =   6120
      TabIndex        =   72
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton CmdAdd 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Ajouter"
      Height          =   495
      Left            =   4680
      TabIndex        =   71
      Top             =   6000
      Width           =   1335
   End
   Begin VB.TextBox txtComment 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   25
      Top             =   1920
      Width           =   4575
   End
   Begin VB.ComboBox cmbFabricant 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1560
      TabIndex        =   10
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
      TabIndex        =   8
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtDescriptionEN 
      BackColor       =   &H00FFFFFF&
      DataField       =   "DESCR_EN"
      DataSource      =   "DatCat1"
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   21
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
      MaxLength       =   100
      TabIndex        =   14
      Top             =   960
      Width           =   4575
   End
   Begin VB.TextBox txtNoItem 
      DataField       =   "PIECE"
      DataSource      =   "DatCat1"
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   1800
      Width           =   1932
   End
   Begin VB.CommandButton CmdAnul 
      Caption         =   "A&nnuler"
      Height          =   495
      Left            =   6120
      TabIndex        =   73
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton CmdEnr 
      Caption         =   "&Enregistre"
      Default         =   -1  'True
      Height          =   495
      Left            =   4680
      TabIndex        =   70
      Top             =   6000
      Width           =   1335
   End
   Begin VB.TextBox txtCategorie 
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   4215
   End
   Begin VB.CommandButton cmdRechercherPiece 
      Height          =   375
      Left            =   3600
      Picture         =   "FrmCatalogueMec.frx":107384
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1800
      Width           =   375
   End
   Begin VB.ComboBox cmbDescriptionFR 
      Height          =   315
      Left            =   5280
      TabIndex        =   13
      Top             =   960
      Width           =   4575
   End
   Begin VB.CommandButton cmdRechercherManufact 
      Height          =   375
      Left            =   3600
      Picture         =   "FrmCatalogueMec.frx":1076C6
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   960
      Width           =   375
   End
   Begin VB.CheckBox chkInventaire 
      BackColor       =   &H00000000&
      Caption         =   "Présent dans l'inventaire"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5280
      TabIndex        =   26
      Top             =   2280
      Width           =   1335
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
      Height          =   975
      Left            =   6840
      TabIndex        =   28
      Top             =   2280
      Width           =   3495
      Begin VB.TextBox txtQuantiteCommande 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         TabIndex        =   36
         Top             =   600
         Width           =   495
      End
      Begin VB.CheckBox chkMinimum 
         BackColor       =   &H00000000&
         Caption         =   "Minimum :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1800
         TabIndex        =   31
         Top             =   240
         Width           =   1060
      End
      Begin VB.TextBox txtQuantiteStock 
         Height          =   285
         Left            =   1200
         TabIndex        =   33
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtQuantiteMinimum 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         TabIndex        =   32
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtQuantitéBoite 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   30
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox chkBoite 
         BackColor       =   &H00000000&
         Caption         =   "Par Boîte :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "À commander :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1800
         TabIndex        =   35
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.ComboBox cmbLocalisation 
      Height          =   315
      Left            =   5160
      TabIndex        =   39
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox txtLocalisation 
      Height          =   285
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   40
      Text            =   "Text1"
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton cmdRechercheAchat 
      Caption         =   "Recherche dans achats"
      Height          =   495
      Left            =   2520
      TabIndex        =   80
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Localisation :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5160
      TabIndex        =   27
      Top             =   2760
      Width           =   975
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
      Left            =   3480
      TabIndex        =   1
      Top             =   120
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
      Width           =   1695
   End
   Begin VB.Label Label1 
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
      Left            =   4080
      TabIndex        =   22
      Top             =   1920
      Width           =   1215
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
      TabIndex        =   20
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
      TabIndex        =   12
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
      TabIndex        =   7
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
      TabIndex        =   16
      Top             =   1800
      Width           =   1095
   End
End
Attribute VB_Name = "FrmCatalogueMec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Numéros de colonne du ListView pour la recherche par description
Private Const I_COL_DES_DESCR_FR    As Integer = 0
Private Const I_COL_DES_DESCR_EN    As Integer = 1
Private Const I_COL_DES_FABRICANT   As Integer = 2
Private Const I_COL_DES_PIECE       As Integer = 3

'Numéros de colonne du ListView pour la recherche par pièce
Private Const I_COL_PIECE_PIECE     As Integer = 0
Private Const I_COL_PIECE_FABRICANT As Integer = 1
Private Const I_COL_PIECE_DESCR_FR  As Integer = 2
Private Const I_COL_PIECE_DESCR_EN  As Integer = 3

'Numéros de colonne du ListView pour la recherche par manufacturier
Private Const I_COL_MAN_FABRICANT   As Integer = 0
Private Const I_COL_MAN_PIECE       As Integer = 1
Private Const I_COL_MAN_DESCR_FR    As Integer = 2
Private Const I_COL_MAN_DESCR_EN    As Integer = 3

'Numéros de colonne du ListView pour les fournisseurs
Private Const I_COL_FRS_FRS         As Integer = 0
Private Const I_COL_FRS_PERS_RESS   As Integer = 1
Private Const I_COL_FRS_DATE        As Integer = 2
Private Const I_COL_FRS_ENTRER_PAR  As Integer = 3
Private Const I_COL_FRS_VALIDE      As Integer = 4
Private Const I_COL_FRS_PRIX_LIST   As Integer = 5
Private Const I_COL_FRS_ESCOMPTE    As Integer = 6
Private Const I_COL_FRS_PRIX_NET    As Integer = 7
Private Const I_COL_FRS_PRIX_SP     As Integer = 8
Private Const I_COL_FRS_QUOTER      As Integer = 9

'Numéros de colonne du ListView pour la recherche dans les jobs
Private Const I_COL_JOB_NUMERO      As Integer = 0
Private Const I_COL_JOB_QUANTITE    As Integer = 1

Private Const I_COL_ACHAT_NUMERO    As Integer = 0
Private Const I_COL_ACHAT_QUANTITE  As Integer = 1

Public Enum enumModeCatalogueMec
  MODE_AJOUT_MODIF_MEC = 0
  MODE_INACTIF = 1
  MODE_AJOUT_MODIF_FRS = 2
End Enum

Public m_eDemande            As enumModeDemande
Public m_bDemandeAnnuler     As Boolean
Public m_bAjout              As Boolean
Public m_sCategorieCopie     As String
Public m_bAnnulerCopie       As Boolean

Private m_bRempliManuel      As Boolean
Private m_sNoItem            As String
Private m_collPieceDescFR    As Collection
Private m_bBloqueDescription As Boolean
Private m_eMode              As enumModeCatalogueMec

Public m_sSelectCategorie    As String
Public m_sSelectFabricant    As String
Public m_sSelectNoItem       As String

'Pour pouvoir comparer la quantité stock avant et après une modification
'pour savoir que c'est de l'ajustement d'inventaire
Private m_sQteStockAvant     As String

Private m_bCopiePiece        As Boolean
Public lastentry             As Boolean
'utilisé pour créer la condition pour les recordsets si on choisi tous les fabricant
Public sChoisirTous          As String


Public Sub ViderChamps_frs()

5       On Error GoTo AfficherErreur

        'Enlever la sélection dans le combo
10      cmbfrs.ListIndex = -1

        'Vide les champs pieces
15      txtPrixSpecial.Text = vbNullString
20      cmbPersRess.ListIndex = -1
25      txtPrixList.Text = vbNullString
30      mskEscompte.Text = vbNullString
35      txtPrixNet.Text = vbNullString
40      mskValide.Text = vbNullString
  
        'Enlève le check
45      chkquoter.Value = vbUnchecked
50      optCAN.Value = True

55      Exit Sub

AfficherErreur:

60      woups "frmCatalogueMec", "ViderChamps_frs", Err, Erl
End Sub

Public Sub ViderChamps_piece()

5       On Error GoTo AfficherErreur

        'Vide les champs pieces
10      txtNoItemGRB.Text = vbNullString
15      txtDescriptionEN.Text = vbNullString
20      txtComment.Text = vbNullString
25      txtQuantitéBoite.Text = vbNullString
30      txtQuantiteCommande.Text = vbNullString
35      txtQuantiteMinimum.Text = vbNullString
40      txtQuantiteStock.Text = vbNullString
45      txtLocalisation.Text = vbNullString

50      cmbLocalisation.ListIndex = -1

55      chkBoite.Value = vbUnchecked
60      chkInventaire.Value = vbUnchecked
65      chkMinimum.Value = vbUnchecked

70      Exit Sub

AfficherErreur:

75      woups "frmCatalogueMec", "ViderChamps_piece", Err, Erl
End Sub

Public Sub BarrerChamps_piece(ByVal bLocked As Boolean)

5       On Error GoTo AfficherErreur

        'Barre les champs
10      txtNoItem.Locked = bLocked
15      txtNoItemGRB.Locked = bLocked
20      txtDescriptionEN.Locked = bLocked
25      txtDescriptionFR.Locked = bLocked
30      txtComment.Locked = bLocked
35      frafournisseur.Enabled = bLocked
40      chkInventaire.Enabled = Not bLocked

45      If chkInventaire.Enabled = True Then
50        If chkInventaire.Value = vbChecked Then
55          fraQuantité.Enabled = True
60        Else
65          fraQuantité.Enabled = False
70        End If
75      Else
80        fraQuantité.Enabled = False
85      End If

90      Exit Sub

AfficherErreur:

95      woups "frmCatalogueMec", "BarrerChamps_piece", Err, Erl
End Sub

Public Sub MontrerControles(ByVal eMode As enumModeCatalogueMec)

5       On Error GoTo AfficherErreur

        'Mets des champs visible et d'autres invisible
10      Dim bTable          As Boolean
15      Dim bFabricant      As Boolean
20      Dim bNoItem         As Boolean
25      Dim bLocalisation   As Boolean
30      Dim bCmdAddFRS      As Boolean
35      Dim bCmdModifFRS    As Boolean
40      Dim bCmdSuppFRS     As Boolean
45      Dim bCmdEnrFRS      As Boolean
50      Dim bCmdAnnulFRS    As Boolean
55      Dim bCmdAdd         As Boolean
60      Dim bCmdModif       As Boolean
65      Dim bCmdSupp        As Boolean
70      Dim bCmdFermer      As Boolean
75      Dim bCmdEnr         As Boolean
80      Dim bCmdAnnul       As Boolean
85      Dim bFraFRS         As Boolean
90      Dim bLvwFRS         As Boolean
95      Dim bCmdSearchMan   As Boolean
100     Dim bCmdSearchPiece As Boolean
105     Dim bCmdSearchDescr As Boolean
110     Dim bCmdDemande     As Boolean
115     Dim bCategorie      As Boolean
120     Dim bCmbDescFR      As Boolean
125     Dim bCopier         As Boolean
130     Dim bChangerCat     As Boolean
135     Dim bInventaire     As Boolean

140     m_eMode = eMode

145     Select Case eMode
          Case MODE_INACTIF:
150         bTable = True
155         bFabricant = True
160         bNoItem = True
165         bCmdAddFRS = True
170         bCmdModifFRS = True
175         bCmdSuppFRS = True
180         bCmdAdd = True
185         bCmdModif = True
190         bCmdSupp = True
195         bCmdFermer = True
200         bFraFRS = True
205         bLvwFRS = True
210         bCmdSearchMan = True
215         bCmdSearchPiece = True
220         bCmdSearchDescr = True
225         bCmdDemande = True
230         bCategorie = True
235         bCopier = True
240         bCmbDescFR = True
245         bInventaire = True
      
          Case MODE_AJOUT_MODIF_MEC:
250         bCmdAddFRS = True
255         bCmdModifFRS = True
260         bCmdSuppFRS = True
265         bCmdEnr = True
            bFabricant = True 'GLL 2017-10-10
            txtFabricant.Enabled = True 'GLL 2017-10-10
270         bCmdAnnul = True
275         bLvwFRS = True
280         bCmdSearchDescr = True
285         bLocalisation = True
290         bChangerCat = True
              
          Case MODE_AJOUT_MODIF_FRS:
295         bCmdEnrFRS = True
300         bCmdAnnulFRS = True
305         bFraFRS = True
310     End Select
  
315     cmbCategorie.Visible = bTable
320     txtCategorie.Visible = Not bTable
  
325     cmbDescriptionFR.Visible = bCmbDescFR
330     txtDescriptionFR.Visible = Not bCmbDescFR

335     cmbFabricant.Visible = bFabricant
340     txtFabricant.Visible = bFabricant
  
345     cmbNoItem.Visible = bNoItem
350     txtNoItem.Visible = Not bNoItem
  
355     cmbLocalisation.Visible = bLocalisation
360     txtLocalisation.Visible = Not bLocalisation

365     frafournisseur.Enabled = bFraFRS
  
370     lvwfournisseur.Visible = bLvwFRS
  
375     cmdAddFrs.Visible = bCmdAddFRS
380     cmdModifFrs.Visible = bCmdModifFRS
385     cmdSuppFrs.Visible = bCmdSuppFRS
390     cmdEnrFrs.Visible = bCmdEnrFRS
395     cmdAnnulFrs.Visible = bCmdAnnulFRS
400     CmdAdd.Visible = bCmdAdd
405     CmdModif.Visible = bCmdModif
410     CmdSupp.Visible = bCmdSupp
415     CmdFerme.Visible = bCmdFermer
420     CmdEnr.Visible = bCmdEnr
425     CmdAnul.Visible = bCmdAnnul
430     cmdDemande.Visible = bCmdDemande
435     cmdCopier.Visible = bCopier
440     cmdRechercheDescrFR.Enabled = bCmdSearchDescr
445     cmdRechercherPiece.Enabled = bCmdSearchPiece
450     cmdRechercherManufact.Enabled = bCmdSearchMan
455     cmdChangerCategorie.Visible = bChangerCat
460     cmdRechercheInventaire.Visible = bInventaire

465     lblDevise1.Visible = False
470     txtTauxChange.Visible = False
475     lblDevise2.Visible = False

480     Exit Sub

AfficherErreur:

485     woups "frmCatalogueMec", "MontrerControles", Err, Erl
End Sub

Private Sub RemplirComboPersRess()

5       On Error GoTo AfficherErreur

10      Dim rstContact    As ADODB.Recordset
15      Dim rstContactFRS As ADODB.Recordset
    
20      Call cmbPersRess.Clear

25      Set rstContactFRS = New ADODB.Recordset
30      Set rstContact = New ADODB.Recordset
    
35      Call rstContactFRS.Open("SELECT * FROM GRB_ContactFRS WHERE NoFRS = " & cmbfrs.ItemData(cmbfrs.ListIndex), g_connData, adOpenDynamic, adLockOptimistic)
  
40      Do While Not rstContactFRS.EOF
45        Call rstContact.Open("SELECT NomContact, IDContact FROM GRB_Contact WHERE IDContact = " & rstContactFRS.Fields("NoContact"), g_connData, adOpenDynamic, adLockOptimistic)

50        If Not rstContact.EOF Then
55          Call cmbPersRess.AddItem(rstContact.Fields("NomContact"))
        
60          cmbPersRess.ItemData(cmbPersRess.newIndex) = rstContact.Fields("IDContact")
65        End If
      
70        Call rstContact.Close

75        Call rstContactFRS.MoveNext
80      Loop
  
85      Call rstContactFRS.Close
90      Set rstContactFRS = Nothing
  
95      If cmbPersRess.ListCount = 0 Then
100       Call rstContact.Open("SELECT NomContact, IDContact FROM GRB_Contact WHERE Supprimé = False ORDER BY NomContact", g_connData, adOpenDynamic, adLockOptimistic)
    
105       Do While Not rstContact.EOF
110         Call cmbPersRess.AddItem(rstContact.Fields("NomContact"))
      
115         cmbPersRess.ItemData(cmbPersRess.newIndex) = rstContact.Fields("IDContact")
    
120         Call rstContact.MoveNext
125       Loop
    
130       Call rstContact.Close
135     End If

140       Set rstContact = Nothing

145     Exit Sub

AfficherErreur:

150     woups "frmCatalogueMec", "RemplirComboPersRess", Err, Erl
End Sub

Private Sub chkBoite_Click()
  
5       On Error GoTo AfficherErreur

10      If m_eMode = MODE_AJOUT_MODIF_MEC Then
15        If chkBoite.Value = vbChecked Then
20          txtQuantitéBoite.Enabled = True
25        Else
30          txtQuantitéBoite.Enabled = False
35        End If
40      End If

45      Exit Sub

AfficherErreur:
  
50      woups "frmCatalogueMec", "chkBoite_Click", Err, Erl
End Sub

Private Sub chkInventaire_Click()

5       On Error GoTo AfficherErreur

10      If m_eMode = MODE_AJOUT_MODIF_MEC Then
15        If chkInventaire.Value = vbChecked Then
20          fraQuantité.Enabled = True
25          cmbLocalisation.Enabled = True
30        Else
35          fraQuantité.Enabled = False
40          cmbLocalisation.Enabled = False
45        End If
50      End If

55      Exit Sub

AfficherErreur:

60      woups "frmCatalogueMec", "chkInventaire_Click", Err, Erl
End Sub

Private Sub chkMinimum_Click()

5       On Error GoTo AfficherErreur

10      If m_eMode = MODE_AJOUT_MODIF_MEC Then
15        If chkMinimum.Value = vbChecked Then
20          txtQuantiteMinimum.Enabled = True
25          txtQuantiteCommande.Enabled = True
30        Else
35          txtQuantiteMinimum.Enabled = False
40          txtQuantiteCommande.Enabled = False
45        End If
50      End If

55      Exit Sub

AfficherErreur:

60      woups "frmCatalogueMec", "chkMinimum_Click", Err, Erl
End Sub

Private Sub cmbDescriptionFR_Click()

5       On Error GoTo AfficherErreur

10      Dim rstCatMec As ADODB.Recordset
15      Dim sNoItem   As String
20      Dim iCompteur As Integer

25      txtDescriptionFR.Text = cmbDescriptionFR.Text
  
30      If m_bBloqueDescription = False Then
35        For iCompteur = 0 To cmbNoItem.ListCount - 1
40          If cmbNoItem.LIST(iCompteur) = m_collPieceDescFR(cmbDescriptionFR.ListIndex + 1) Then
45            cmbNoItem.ListIndex = iCompteur

50            Exit For
55          End If
60        Next
65      End If

70      Exit Sub

AfficherErreur:

75      woups "frmCatalogueMec", "cmbDescriptionFR_Click", Err, Erl
End Sub

Private Sub cmbfrs_Click()

5       On Error GoTo AfficherErreur

10      If cmbfrs.ListIndex <> -1 Then
15        cmbfrs.Tag = cmbfrs.ItemData(cmbfrs.ListIndex)
   
20        Call RemplirComboPersRess
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmCatalogueMec", "cmbfrs_Click", Err, Erl
End Sub

Private Sub cmbLocalisation_Click()

5       On Error GoTo AfficherErreur

10      txtLocalisation.Text = cmbLocalisation.Text

15      Exit Sub

AfficherErreur:

20      woups "frmCatalogueMec", "cmbLocalisation_Click", Err, Erl
End Sub

Private Sub CmdAdd_Click()

5       On Error GoTo AfficherErreur

        'montre le dialogue pour ajouter un item au catalogue
10      Screen.MousePointer = vbHourglass

15      m_bBloqueDescription = True
  
20      Call OuvrirForm(FrmaddItemMec, True)
  
25      m_bBloqueDescription = False
  
30      Screen.MousePointer = vbDefault

35      Exit Sub

AfficherErreur:

40      woups "frmCatalogueMec", "CmdAdd_Click", Err, Erl
End Sub

Private Sub cmdAddFrs_Click()

5       On Error GoTo AfficherErreur

        'ajoute un fournisseur pour la piece
10      If cmbNoItem.ListCount > 0 Then
15        m_bAjout = True

20        Call BarrerChamps_piece(True)

25        Call ViderChamps_frs

30        Call cmbfrs.SetFocus

35        Call MontrerControles(MODE_AJOUT_MODIF_FRS)
  
          'affiche drapeau
40        Call AfficherDrapeau
45      End If

50      Exit Sub

AfficherErreur:

55      woups "frmCatalogueMec", "cmdAddFrs_Click", Err, Erl
End Sub

Private Sub cmdAnnulFrs_Click()

5       On Error GoTo AfficherErreur

10      Call MontrerControles(MODE_INACTIF)

15      Exit Sub

AfficherErreur:

20      woups "frmCatalogueMec", "cmdAnnulFrs_Click", Err, Erl
End Sub

Private Sub CmdAnul_Click()

5       On Error GoTo AfficherErreur

10      txtPrixNet.Enabled = True
15      txtPrixSpecial.Enabled = True

20      m_bBloqueDescription = True

25      Call AfficherItem
        txtFabricant.Top = 1320 'GLL 2017-10-10
        cmbFabricant.Visible = True 'GLL 2017-10-10
        
30      m_bBloqueDescription = False
  
35      m_bCopiePiece = False
  
        'on cache les combos
40      cmbFabricant.Visible = False
45      cmbNoItem.Visible = False

        'on retablis les boutons
50      Call MontrerControles(MODE_INACTIF)
55      Call BarrerChamps_piece(True)

60      m_sQteStockAvant = ""

65      Exit Sub

AfficherErreur:

70      woups "frmCatalogueMec", "CmdAnul_Click", Err, Erl
End Sub

Private Sub EnregistrerItem()

5       On Error GoTo AfficherErreur

        'Enregistrement de l'item dans la BD
10      Dim rstItem        As ADODB.Recordset
15      Dim rstItemFRS     As ADODB.Recordset
20      Dim rstItemFRSDest As ADODB.Recordset
25      Dim rstVerif       As ADODB.Recordset
30      Dim rstInventaire  As ADODB.Recordset
35      Dim rstInvModif    As ADODB.Recordset
40      Dim sNomFab        As String
45      Dim sNoPiece       As String
50      Dim iCompteur      As Integer
55      Dim sPieceModif    As String
60      Dim sLettre        As String
  
65      sNomFab = txtFabricant.Text
70      sNoPiece = txtNoItem.Text
  
75      If m_bCopiePiece = True Or (m_bCopiePiece = False And (UCase(sNoPiece) <> UCase(m_sNoItem))) Then
80        Set rstVerif = New ADODB.Recordset

85        Call rstVerif.Open("SELECT * FROM GRB_CatalogueMec WHERE PIECE = '" & Replace(sNoPiece, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
      
90        If Not rstVerif.EOF Then
95          Call MsgBox("Le numéro " & sNoPiece & " existe déjà!", vbOKOnly, "Erreur")
                
100         Call rstVerif.Close
105         Set rstVerif = Nothing
        
110         Exit Sub
115       End If
      
120       Call rstVerif.Close
125       Set rstVerif = Nothing
130     End If
  
135     If txtFabricant.Text = vbNullString Or txtNoItem.Text = vbNullString Or txtDescriptionFR.Text = vbNullString Then
140       Call MsgBox("Les champs Manufacturier, Pièce et Desc. FR doivent être remplis!", vbOKOnly, "Erreur")
          
145       Exit Sub
150     End If
    
        'Sinon, j'ouvre un recordset contenant le no d'item
155     Set rstItem = New ADODB.Recordset
        
160     Call rstItem.Open("SELECT * FROM GRB_CatalogueMec WHERE PIECE = '" & Replace(m_sNoItem, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
  
        'enregistre le nopiece dans la table distributeur si pas vide
165     Set rstItemFRS = New ADODB.Recordset
        
170     Call rstItemFRS.Open("SELECT * FROM GRB_PiecesFRS WHERE PIECE = '" & Replace(rstItem.Fields("PIECE"), "'", "''") & "' AND Type = 'M'", g_connData, adOpenDynamic, adLockOptimistic)
    
175     If m_bCopiePiece = False Then
180       Do While Not rstItemFRS.EOF
185         rstItemFRS.Fields("PIECE") = txtNoItem.Text
     
190         Call rstItemFRS.Update
      
195         Call rstItemFRS.MoveNext
200       Loop
205     Else
210       Set rstItemFRSDest = New ADODB.Recordset

215       Call rstItemFRSDest.Open("SELECT * FROM GRB_PiecesFRS", g_connData, adOpenDynamic, adLockOptimistic)

220       Do While Not rstItemFRS.EOF
225         Call rstItemFRSDest.AddNew

230         rstItemFRSDest.Fields("IDFRS") = rstItemFRS.Fields("IDFRS")
235         rstItemFRSDest.Fields("PIECE") = sNoPiece
240         rstItemFRSDest.Fields("PRIX_SP") = rstItemFRS.Fields("PRIX_SP")
245         rstItemFRSDest.Fields("PERS_RESS") = rstItemFRS.Fields("PERS_RESS")
250         rstItemFRSDest.Fields("PRIX_LIST") = rstItemFRS.Fields("PRIX_LIST")
255         rstItemFRSDest.Fields("ESCOMPTE") = rstItemFRS.Fields("ESCOMPTE")
260         rstItemFRSDest.Fields("PRIX_NET") = rstItemFRS.Fields("PRIX_NET")
265         rstItemFRSDest.Fields("DATE") = rstItemFRS.Fields("DATE")
270         rstItemFRSDest.Fields("ENTRER_PAR") = rstItemFRS.Fields("ENTRER_PAR")
275         rstItemFRSDest.Fields("VALIDE") = rstItemFRS.Fields("VALIDE")
280         rstItemFRSDest.Fields("QUOTER") = rstItemFRS.Fields("QUOTER")
285         rstItemFRSDest.Fields("DeviseMonétaire") = rstItemFRS.Fields("DeviseMonétaire")
290         rstItemFRSDest.Fields("Type") = rstItemFRS.Fields("Type")

295         Call rstItemFRSDest.Update

300         Call rstItemFRS.MoveNext
305       Loop

310       Call rstItemFRSDest.Close
315       Set rstItemFRSDest = Nothing
320     End If

325     Call rstItemFRS.Close
330     Set rstItemFRS = Nothing
   
335     If m_bCopiePiece = True Then
340       Call rstItem.AddNew
345     End If
     
        'Enregistrement des valeurs dans la table catalogue
350     rstItem.Fields("CATEGORIE") = txtCategorie.Text
355     rstItem.Fields("PIECE").Value = sNoPiece

360     For iCompteur = 1 To Len(sNoPiece)
365       sLettre = Mid$(sNoPiece, iCompteur, 1)

370       If (Asc(sLettre) >= 48 And Asc(sLettre) <= 57) Or _
             (Asc(sLettre) >= 65 And Asc(sLettre) <= 90) Or _
             (Asc(sLettre) >= 97 And Asc(sLettre) <= 122) Then
375         sPieceModif = sPieceModif & sLettre
380       End If
385     Next

390     rstItem.Fields("PIECE_MODIF").Value = sPieceModif
395     rstItem.Fields("FABRICANT").Value = sNomFab
400     rstItem.Fields("PIECE_GRB").Value = txtNoItemGRB.Text
405     rstItem.Fields("DESC_EN").Value = txtDescriptionEN.Text
410     rstItem.Fields("DESC_FR").Value = txtDescriptionFR.Text
415     rstItem.Fields("COMMENTAIRE").Value = txtComment.Text
       
420     Call rstItem.Update
  
425     Call rstItem.Close
430     Set rstItem = Nothing

435     If chkInventaire.Value = vbChecked Then
440       Set rstInventaire = New ADODB.Recordset

445       If m_bCopiePiece = True Then
450         Call rstInventaire.Open("SELECT * FROM GRB_InventaireMec WHERE NoItem = '" & Replace(sNoPiece, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
455       Else
460         Call rstInventaire.Open("SELECT * FROM GRB_InventaireMec WHERE NoItem = '" & Replace(m_sNoItem, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
465       End If
  
470       If rstInventaire.EOF Then
475         Call rstInventaire.AddNew
480       End If

485       rstInventaire.Fields("NoItem") = txtNoItem.Text

490       rstInventaire.Fields("Description") = txtDescriptionFR.Text

495       rstInventaire.Fields("Manufacturier") = sNomFab

500       If chkBoite.Value = vbChecked Then
505         rstInventaire.Fields("CommandeParBoite") = True
510         rstInventaire.Fields("QteBoite") = txtQuantitéBoite.Text
515       Else
520         rstInventaire.Fields("CommandeParBoite") = False
525         rstInventaire.Fields("QteBoite") = ""
530       End If

535       Set rstItemFRS = New ADODB.Recordset
          
540       Call rstItemFRS.Open("SELECT * FROM GRB_PiecesFRS WHERE PIECE = '" & Replace(txtNoItem.Text, "'", "''") & "' AND IDFRS = 717", g_connData, adOpenDynamic, adLockOptimistic)
    
545       If rstItemFRS.EOF Then
550         Call rstItemFRS.AddNew

555         rstItemFRS.Fields("PIECE").Value = txtNoItem.Text
560         rstItemFRS.Fields("IDFRS").Value = 717
565         rstItemFRS.Fields("Type").Value = "M"
570         rstItemFRS.Fields("PERS_RESS").Value = Null
575         rstItemFRS.Fields("PRIX_LIST").Value = "0"
580         rstItemFRS.Fields("ESCOMPTE").Value = "0"
585         rstItemFRS.Fields("PRIX_NET").Value = "0"
590         rstItemFRS.Fields("PrixReel").Value = "0"
595         rstItemFRS.Fields("DATE").Value = ConvertDate(Date)
600         rstItemFRS.Fields("ENTRER_PAR").Value = g_sInitiale
605         rstItemFRS.Fields("DeviseMonétaire").Value = "CAN"
   
610         Call rstItemFRS.Update
615       End If

620       If chkBoite.Value = vbChecked Then
625         If IsNumeric(rstItemFRS.Fields("PRIX_LIST")) Then
630           rstInventaire.Fields("Prix Liste") = Round(rstItemFRS.Fields("PRIX_LIST") / txtQuantitéBoite.Text, 6)
635         Else
640           rstInventaire.Fields("Prix Liste") = "0"
645         End If

650         If IsNumeric(rstItemFRS.Fields("ESCOMPTE")) Then
655           rstInventaire.Fields("Escompte") = rstItemFRS.Fields("ESCOMPTE")
660         Else
665           rstInventaire.Fields("Escompte") = "0"
670         End If

675         If IsNumeric(rstItemFRS.Fields("PRIX_NET")) Then
680           rstInventaire.Fields("Prix net") = Round(rstItemFRS.Fields("PRIX_NET") / txtQuantitéBoite.Text, 6)
685         Else
690           rstInventaire.Fields("Prix net") = "0"
695         End If
700       Else
705         rstInventaire.Fields("Prix Liste") = rstItemFRS.Fields("PRIX_LIST")
710         rstInventaire.Fields("Escompte") = rstItemFRS.Fields("ESCOMPTE")
715         rstInventaire.Fields("Prix net") = rstItemFRS.Fields("PRIX_NET")
720       End If

725       Call rstItemFRS.Close
730       Set rstItemFRS = Nothing

735       rstInventaire.Fields("Commentaires") = txtComment.Text

740       rstInventaire.Fields("Localisation") = cmbLocalisation.Text

745       If Trim$(txtQuantiteStock.Text) <> "" Then
750         rstInventaire.Fields("QuantitéStock") = txtQuantiteStock.Text
755       Else
760         rstInventaire.Fields("QuantitéStock") = "0"
765       End If

770       If chkMinimum.Value = vbChecked Then
775         rstInventaire.Fields("Minimum") = True

780         If Trim$(txtQuantiteMinimum.Text) <> "" Then
785           rstInventaire.Fields("QuantitéMinimum") = txtQuantiteMinimum.Text
790         Else
795           rstInventaire.Fields("QuantitéMinimum") = "0"
800         End If

805         If Trim$(txtQuantiteCommande.Text) = True Then
810           rstInventaire.Fields("Commande") = txtQuantiteCommande.Text
815         Else
820           rstInventaire.Fields("Commande") = "0"
825         End If
830       Else
835         rstInventaire.Fields("Minimum") = False
840         rstInventaire.Fields("QuantitéMinimum") = ""
845         rstInventaire.Fields("Commande") = ""
850       End If

855       Call rstInventaire.Update

860       Call rstInventaire.Close
865       Set rstInventaire = Nothing
870     Else
875       If m_bCopiePiece = True Then
880         Call g_connData.Execute("DELETE * FROM GRB_InventaireMec WHERE NoItem = '" & Replace(sNoPiece, "'", "''") & "'")
885       Else
890         Call g_connData.Execute("DELETE * FROM GRB_InventaireMec WHERE NoItem = '" & Replace(m_sNoItem, "'", "''") & "'")
895       End If
900     End If

905     If m_bCopiePiece = False Then
910       If txtQuantiteStock.Text <> m_sQteStockAvant Or ((m_sQteStockAvant <> "" And m_sQteStockAvant <> "0") And chkInventaire.Value = vbUnchecked) Then
915         If m_sQteStockAvant = "" Then
920           m_sQteStockAvant = "0"
925         End If

930         If Not IsNumeric(txtQuantiteStock.Text) Then
935           txtQuantiteStock.Text = "0"
940         End If

945         Set rstInvModif = New ADODB.Recordset

950         Call rstInvModif.Open("SELECT * FROM GRB_InventaireMecModif", g_connData, adOpenDynamic, adLockOptimistic)

955         Call rstInvModif.AddNew

960         rstInvModif.Fields("Date") = ConvertDate(Date)
965         rstInvModif.Fields("IDProjet") = InputBox("Précisez l'ajustement d'inventaire")
970         rstInvModif.Fields("NoItem") = txtNoItem.Text

975         If chkInventaire.Value = vbChecked Then
980           rstInvModif.Fields("Quantité") = CDbl(txtQuantiteStock.Text) - CDbl(m_sQteStockAvant)
985         Else
990           rstInvModif.Fields("Quantité") = 0 - CDbl(m_sQteStockAvant)
995         End If

1000        rstInvModif.Fields("User") = g_sInitiale

1005        Call rstInvModif.Update

1010        Call rstInvModif.Close
1015        Set rstInvModif = Nothing
1020      End If
1025    End If
    
1030    If (UCase(sNoPiece) <> UCase(m_sNoItem)) And m_bCopiePiece = False Then
1035      Call ModifierNoItem(m_sNoItem, sNoPiece)
1040    End If
    
1045    m_sQteStockAvant = ""
    
1050    m_bRempliManuel = True
  
1055    m_sSelectNoItem = sNoPiece
1060    m_sSelectFabricant = sNomFab

1065    Call RemplirComboLocalisation

1070    Call RemplirComboFabricant
  
        'Rétablir les buttons
1075    Call MontrerControles(MODE_INACTIF)
  
1080    Call BarrerChamps_piece(True)

1085    Exit Sub

AfficherErreur:

1090    woups "frmCatalogueMec", "EnregistrerItem", Err, Erl
End Sub

Private Sub cmdChangerCategorie_Click()

5       On Error GoTo AfficherErreur

10      Dim rstPiece As Recordset

15      Call frmChoixCategorie.Afficher(MECANIQUE)

20      If txtCategorie.Text <> m_sCategorieCopie Then
25        If m_bAnnulerCopie = False Then
30          Set rstPiece = New ADODB.Recordset

35          Call rstPiece.Open("SELECT * FROM GRB_CatalogueMec WHERE PIECE = '" & Replace(cmbNoItem.Text, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)

40          rstPiece.Fields("CATEGORIE") = m_sCategorieCopie

45          Call rstPiece.Update

50          Call rstPiece.Close
55          Set rstPiece = Nothing

60          Call ViderChamps_piece

65          m_sSelectFabricant = txtFabricant.Text

70          Call RemplirComboFabricant

75          Call MontrerControles(MODE_INACTIF)

80          Call BarrerChamps_piece(True)
85        End If
90      End If

95      Exit Sub

AfficherErreur:

100     woups "frmCatalogueMec", "cmdChangerCategorie_Click", Err, Erl
End Sub

Private Sub cmdCopier_Click()

5       On Error GoTo AfficherErreur

10      m_bCopiePiece = True

15      Call CmdModif_Click

20      chkInventaire.Value = vbUnchecked
25      chkBoite.Value = vbUnchecked
30      chkMinimum.Value = vbUnchecked

35      txtQuantitéBoite.Text = ""
40      txtQuantiteStock.Text = ""
45      txtQuantiteMinimum.Text = ""
50      txtQuantiteCommande.Text = ""
55      cmbLocalisation.Text = ""

60      Exit Sub

AfficherErreur:

65      woups "frmCatalogueMec", "cmdCopier_Click", Err, Erl
End Sub

Private Sub cmdDemande_Click()

5       On Error GoTo AfficherErreur

10      Call dlgDemandePrix.Afficher(Me)
  
15      If m_bDemandeAnnuler = False Then
20        Call frmChoixDemande.Afficher(MECANIQUE, m_eDemande)
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmCatalogueMec", "cmdDemande_Click", Err, Erl
End Sub

Private Sub CmdEnr_Click()

5       On Error GoTo AfficherErreur

10      Dim iCompteur  As Integer
15      Dim bContinuer As Boolean

        'Enregistrement d'un item dans la BD
        txtFabricant.Top = 1320
        cmbFabricant.Visible = False
20      If (UCase(txtNoItem.Text) <> UCase(m_sNoItem)) And m_bCopiePiece = False Then
25        If MsgBox("Le numéro de pièce sera modifié dans toutes les soumissions, les projets et les achats. " & vbNewLine & _
                    "Voulez-vous continuer ?", vbYesNo) = vbYes Then
30          bContinuer = True
35        Else
40          bContinuer = False
45        End If
50      Else
55        bContinuer = True
60      End If

65      If bContinuer = True Then
70        Call EnregistrerItem

75        If m_eMode = MODE_INACTIF Then
80          m_bCopiePiece = False
85        End If
  
90        Call RemplirComboDescription

95        m_bBloqueDescription = True

100       For iCompteur = 0 To cmbDescriptionFR.ListCount - 1
105         If cmbDescriptionFR.LIST(iCompteur) = txtDescriptionFR.Text Then
110           cmbDescriptionFR.ListIndex = iCompteur

115           Exit For
120         End If
125       Next

130       m_bBloqueDescription = False
135     End If

140     Exit Sub

AfficherErreur:

145     woups "frmCatalogueMec", "CmdEnr_Click", Err, Erl
End Sub

Private Sub ModifierNoItem(ByVal sAncienNoItem As String, ByVal sNouveauNoItem As String)
  
5       On Error GoTo AfficherErreur

10      Dim iRecordProjet As Integer
15      Dim iRecordSoum   As Integer
20      Dim iRecordAchat  As Integer

25      Call g_connData.Execute("UPDATE GRB_Projet_Pieces SET NumItem = '" & Replace(sNouveauNoItem, "'", "''") & "' WHERE NumItem = '" & Replace(sAncienNoItem, "'", "''") & "' AND Type = 'M'", iRecordProjet)
30      Call g_connData.Execute("UPDATE GRB_Soumission_Pieces SET NumItem = '" & Replace(sNouveauNoItem, "'", "''") & "' WHERE NumItem = '" & Replace(sAncienNoItem, "'", "''") & "' AND Type = 'M'", iRecordSoum)

35      Call g_connData.Execute("UPDATE GRB_Achat_Pieces SET PIECE = '" & Replace(sNouveauNoItem, "'", "''") & "' WHERE PIECE = '" & Replace(sAncienNoItem, "'", "''") & "' AND Left(IDAchat, 1) = 'M'", iRecordAchat)

40      Call g_connData.Execute("UPDATE GRB_InventaireMecModif SET NoItem = '" & Replace(sNouveauNoItem, "'", "''") & "' WHERE NoItem = '" & Replace(sAncienNoItem, "'", "''") & "'")

45      Call MsgBox("Numéros de pièces modifiés" & vbNewLine & _
                    vbNewLine & _
                    "Projets : " & iRecordProjet & vbNewLine & _
                    "Soumissions : " & iRecordSoum & vbNewLine & _
                    "Achats : " & iRecordAchat, vbOKOnly)

50      Exit Sub

AfficherErreur:

55      woups "frmCatalogueMec", "ModifierNoItem", Err, Erl
End Sub

Private Sub cmdEnrFrs_Click()

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer

        'Enregistrement d'un Item dans la BD
15      If Trim$(txtPrixList.Text) = "" And Trim$(mskEscompte.Text) = "" And Trim$(txtPrixNet.Text) = "" And Trim$(txtPrixSpecial.Text) = "" Then
20        txtPrixList.Text = "0"
25        mskEscompte.Text = "0"
30        txtPrixNet.Text = "0"
35      End If
       
40      If Trim$(txtPrixList.Text) = vbNullString Then
45        If Trim$(txtPrixSpecial.Text) = vbNullString Then
50          Call MsgBox("Vous devez remplir le prix listé!", vbOKOnly, "Erreur")
    
55          Exit Sub
60        End If
65      End If
  
70      If Trim$(txtPrixNet.Text) = vbNullString And Trim$(txtPrixSpecial.Text) = vbNullString Then
75        Call MsgBox("Vous devez remplir le prix net ou le prix spécial!", vbOKOnly, "Erreur")
    
80        Exit Sub
85      End If

90      If optUSA.Value = True Or optSpain.Value = True Then
95        If Trim$(txtTauxChange.Text) <> vbNullString Then
100         If Not IsNumeric(txtTauxChange.Text) Then
105           Call MsgBox("Le taux de change doit être numérique!", vbOKOnly, "Erreur")

110           Exit Sub
115         End If
120       Else
125         Call MsgBox("Le taux de change ne doit pas être vide!", vbOKOnly, "Erreur")

130         Exit Sub
135       End If
140     End If

145     If (Trim$(txtPrixNet.Text) <> Trim$(txtPrixList.Text)) And Trim$(txtPrixSpecial.Text) = vbNullString Then
150       Call CalculerPrixNet
155     End If
 
160     If cmbfrs.ListIndex > -1 Then
165       Call EnregistrerFRS
170     Else
175       Call MsgBox("Vous devez choisir un fournisseur!", vbOKOnly, "Erreur")
180     End If

185     Exit Sub

AfficherErreur:

190     woups "frmCatalogueMec", "cmdEnrFrs_Click", Err, Erl
End Sub

Private Sub EnregistrerFRS()

5       On Error GoTo AfficherErreur

        'Enregistrement de l'item dans la BD
10      Dim rstItemFRS     As ADODB.Recordset
15      Dim rstInv         As ADODB.Recordset
20      Dim rstConfig      As ADODB.Recordset
25      Dim bDistribExiste As Boolean
30      Dim iCompteur      As Integer
      
        'Si le PRIX_SP est monétaire
35      If txtPrixSpecial.Text <> vbNullString Then
40        If Not IsNumeric(txtPrixSpecial.Text) Then
45          Call MsgBox("Le prix spécial est invalide!", vbOKOnly, "Erreur")
       
50          Exit Sub
55        End If
60      End If
   
        'Si le PRIX_NET est monétaire
65      If txtPrixNet.Text <> vbNullString Then
70        If Not IsNumeric(txtPrixNet.Text) Then
75          Call MsgBox("Le prix net est invalide!", vbOKOnly, "Erreur")
      
80          Exit Sub
85        End If
90      End If
    
        'Si le PRIX_LIST est monétaire
95      If txtPrixList.Text <> vbNullString Then
100       If Not IsNumeric(txtPrixList.Text) Then
105         Call MsgBox("Le prix listé est invalide!", vbOKOnly, "Erreur")
      
110         Exit Sub
115       End If
120     End If
    
        'Si la date de validité est valide
125     If Trim$(mskValide.Text) <> vbNullString Then
130       If IsDate(mskValide.Text) = False Then
135         Call MsgBox("La date de validité est invalide!", vbOKOnly, "Erreur")
      
140         Exit Sub
145       End If
150     End If
  
155     bDistribExiste = False
  
160     If m_bAjout = True Then
          'Si le distributeur n'est pas déjà dans le listView
165       If lvwfournisseur.ListItems.count > 0 Then
170         For iCompteur = 1 To lvwfournisseur.ListItems.count
175           If lvwfournisseur.ListItems(iCompteur).Text = cmbfrs.Text Then
180             bDistribExiste = True
        
185             Exit For
190           End If
195         Next
200       End If
  
205       If bDistribExiste = True Then
210         If txtPrixSpecial.Text <> "" Then
215           If lvwfournisseur.ListItems(iCompteur).SubItems(I_COL_FRS_PRIX_SP) <> "" Then
220             Call MsgBox("Ce distributeur est déjà ajouté avec un prix spécial", vbOKOnly, "Erreur")

225             Exit Sub
230           End If
235         Else
240           If lvwfournisseur.ListItems(iCompteur).SubItems(I_COL_FRS_PRIX_NET) <> "" Then
245             Call MsgBox("Ce distributeur est déjà ajouté avec un prix net", vbOKOnly, "Erreur")

250             Exit Sub
255           End If
260         End If
265       End If
270     End If

275     Set rstItemFRS = New ADODB.Recordset

280     If m_bAjout = True Then
285       Call rstItemFRS.Open("SELECT * FROM GRB_PiecesFRS WHERE PIECE = '" & Replace(txtNoItem.Text, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
    
          'si c'est un ajout, j'ouvre un recordset général
290       Call rstItemFRS.AddNew
    
295       m_bAjout = False
300     Else
305       Call rstItemFRS.Open("SELECT * FROM GRB_PiecesFRS WHERE noEnreg = " & lvwfournisseur.SelectedItem.Tag, g_connData, adOpenDynamic, adLockOptimistic)
310     End If
           
        'Enregistrement des valeurs dans la table catalogue
315     rstItemFRS.Fields("PIECE").Value = cmbNoItem.Text
320     rstItemFRS.Fields("IDFRS").Value = cmbfrs.Tag
325     rstItemFRS.Fields("Type").Value = "M"

330     If cmbPersRess.ListIndex > -1 Then
335       rstItemFRS.Fields("PERS_RESS").Value = cmbPersRess.ItemData(cmbPersRess.ListIndex)
340     Else
345       rstItemFRS.Fields("PERS_RESS").Value = Null
350     End If

355     rstItemFRS.Fields("PRIX_LIST").Value = txtPrixList.Text
360     rstItemFRS.Fields("ESCOMPTE").Value = mskEscompte.Text
      
365     If txtPrixSpecial.Text <> vbNullString Or txtPrixNet.Text <> vbNullString Then
370       If txtPrixNet.Text <> vbNullString Then
375         rstItemFRS.Fields("PRIX_NET").Value = txtPrixNet.Text
380         rstItemFRS.Fields("PrixReel").Value = txtPrixNet.Text
385       Else
390         rstItemFRS.Fields("PRIX_NET").Value = vbNullString
395       End If
  
400       If txtPrixSpecial.Text <> vbNullString Then
405         rstItemFRS.Fields("PRIX_SP").Value = txtPrixSpecial.Text
410         rstItemFRS.Fields("PrixReel").Value = txtPrixNet.Text
415       Else
420         rstItemFRS.Fields("PRIX_SP").Value = vbNullString
425       End If
430     End If
        
435     rstItemFRS.Fields("DATE").Value = ConvertDate(Date)
440     rstItemFRS.Fields("VALIDE").Value = mskValide.Text
445     rstItemFRS.Fields("ENTRER_PAR").Value = g_sInitiale
  
450     If chkquoter.Value = 1 Then
455       rstItemFRS.Fields("Quoter").Value = True
460     Else
465       rstItemFRS.Fields("Quoter").Value = False
470     End If
  
475     If optCAN.Value = True Then
480       rstItemFRS.Fields("DeviseMonétaire").Value = "CAN"
485     Else
490       If optUSA.Value = True Then
495         rstItemFRS.Fields("DeviseMonétaire").Value = "USA"
500       Else
505         rstItemFRS.Fields("DeviseMonétaire").Value = "SPA"
510       End If
515     End If
    
520     Call rstItemFRS.Update
   
525     Call rstItemFRS.Close
530     Set rstItemFRS = Nothing

535     If optUSA.Value = True Or optSpain.Value = True Then
540       Set rstConfig = New ADODB.Recordset

545       Call rstConfig.Open("SELECT TauxAmericain, TauxEspagnol FROM GRB_Config", g_connData, adOpenDynamic, adLockOptimistic)

550       If optUSA.Value = True Then
555         rstConfig.Fields("TauxAmericain") = txtTauxChange.Text
560       Else
565         rstConfig.Fields("TauxEspagnol") = txtTauxChange.Text
570       End If

575       Call rstConfig.Update

580       Call rstConfig.Close
585       Set rstConfig = Nothing
590     End If

        'retablir les buttons
595     Call MontrerControles(MODE_INACTIF)

600     If cmbfrs.ItemData(cmbfrs.ListIndex) = 717 Then
605       Set rstInv = New ADODB.Recordset

610       Call rstInv.Open("SELECT * FROM GRB_InventaireMec WHERE NoItem = '" & Replace(txtNoItem.Text, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)

625       If Not rstInv.EOF Then
630         If txtPrixNet.Text <> "" Then
635           If rstInv.Fields("CommandeParBoite") = True Then
640             rstInv.Fields("Prix Liste") = txtPrixList.Text / rstInv.Fields("QteBoite")
645             rstInv.Fields("Escompte") = mskEscompte.Text
650             rstInv.Fields("Prix net") = txtPrixNet.Text / rstInv.Fields("QteBoite")
655           Else
660             rstInv.Fields("Prix Liste") = txtPrixList.Text
665             rstInv.Fields("Escompte") = mskEscompte.Text
670             rstInv.Fields("Prix net") = txtPrixNet.Text
675           End If
680         Else
685           If rstInv.Fields("CommandeParBoite") = True Then
690             rstInv.Fields("Prix Liste") = txtPrixSpecial.Text / rstInv.Fields("QteBoite")
695             rstInv.Fields("Escompte") = ""
700             rstInv.Fields("Prix net") = txtPrixSpecial.Text / rstInv.Fields("QteBoite")
705           Else
710             rstInv.Fields("Prix Liste") = txtPrixSpecial.Text
715             rstInv.Fields("Escompte") = ""
720             rstInv.Fields("Prix net") = txtPrixSpecial.Text
725           End If
730         End If

735         Call rstInv.Update
740       End If

745       Call rstInv.Close
750       Set rstInv = Nothing
755     End If

        'remplis le lister
760     Call RemplirListViewFournisseur

765     Exit Sub

AfficherErreur:

770     woups "frmCatalogueMec", "EnregistrerFRS", Err, Erl
End Sub

Private Sub CmdFerme_Click()

5       On Error GoTo AfficherErreur

        'Fermeture de la fenêtre
10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmCatalogueMec", "CmdFerme_Click", Err, Erl
End Sub

Private Sub CmdModif_Click()

5       On Error GoTo AfficherErreur

        'procedure qui permet de modifier l'enregistrement courant
        'on montre/cache les maskedBox
10      If cmbNoItem.ListCount > 0 Then
          'Copie le contenu textbox dans les maskbox
15        Call MontrerControles(MODE_AJOUT_MODIF_MEC)
          txtFabricant.Top = 960
          cmbFabricant.Visible = False

20        m_sQteStockAvant = txtQuantiteStock.Text
  
25        Call BarrerChamps_piece(False)
30      End If

35      Exit Sub

AfficherErreur:

40      woups "frmCatalogueMec", "CmdModif_Click", Err, Erl
End Sub

Private Sub cmdModifFrs_Click()

5       On Error GoTo AfficherErreur
        'modifie un fournisseur pour la piece
        
10      If lvwfournisseur.ListItems.count > 0 Then
15        Call ModifierFournisseur
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmCatalogueMec", "cmdModifFrs_Click", Err, Erl
End Sub

Private Sub cmdRechercheCategorie_Click()
5
10      Dim rstcatalog As ADODB.Recordset
15      Dim sDescription   As String
20      Dim itmDescription As ListItem
  
25      sDescription = InputBox("Quelle est la catégorie à rechercher")
  
30      If sDescription <> vbNullString Then
35        Call lvwCategorie.ListItems.Clear
  
40        sDescription = Replace(sDescription, "'", "''")
  
45        sDescription = "%" & sDescription & "%"

50        Set rstcatalog = New ADODB.Recordset

55
    Call rstcatalog.Open("SELECT DISTINCT CATEGORIE FROM GRB_CatalogueMec WHERE  Categorie LIKE '" & sDescription & "'  ORDER BY CATEGORIE", g_connData, adOpenDynamic, adLockOptimistic)
60        Do While Not rstcatalog.EOF
65          Set itmDescription = lvwCategorie.ListItems.Add()
        
70          itmDescription.Tag = rstcatalog.Fields("CATEGORIE")
            itmDescription.Text = rstcatalog.Fields("CATEGORIE")

155         Call rstcatalog.MoveNext
160       Loop
    
165       Call rstcatalog.Close
170       Set rstcatalog = Nothing

175       If lvwCategorie.ListItems.count > 0 Then
180         lvwCategorie.Visible = True

185         Call lvwCategorie.SetFocus
190       Else
195         Call MsgBox("Aucun enregistrement trouvé!")
200       End If
205     End If

210     Exit Sub
End Sub

Private Sub cmdRechercheDescrFR_Click()

5       On Error GoTo AfficherErreur

10      Dim rstDescription As ADODB.Recordset
15      Dim sDescription   As String
20      Dim itmDescription As ListItem
  
25      sDescription = InputBox("Quelle est la description à rechercher")
  
30      If sDescription <> vbNullString Then
35        Call lvwDescription.ListItems.Clear
    
40        sDescription = Replace(sDescription, "'", "''")
    
45        sDescription = "%" & sDescription & "%"
           
50        Set rstDescription = New ADODB.Recordset
      
55        Call rstDescription.Open("SELECT * FROM GRB_CatalogueMec WHERE DESC_FR LIKE '" & sDescription & "' OR DESC_EN LIKE '" & sDescription & "' ORDER BY DESC_FR, DESC_EN", g_connData, adOpenDynamic, adLockOptimistic)
      
60        Do While Not rstDescription.EOF
65          Set itmDescription = lvwDescription.ListItems.Add
        
70          itmDescription.Tag = rstDescription.Fields("CATEGORIE")
         
75          If Not IsNull(rstDescription.Fields("DESC_FR")) Then
80            itmDescription.Text = Trim(rstDescription.Fields("DESC_FR"))
85          Else
90            itmDescription.Text = vbNullString
95          End If
      
100         If Not IsNull(rstDescription.Fields("DESC_EN")) Then
105           itmDescription.SubItems(I_COL_DES_DESCR_EN) = Trim(rstDescription.Fields("DESC_EN"))
110         Else
115           itmDescription.SubItems(I_COL_DES_DESCR_EN) = vbNullString
120         End If
          
125         If Not IsNull(rstDescription.Fields("FABRICANT")) Then
130           itmDescription.SubItems(I_COL_DES_FABRICANT) = Trim(rstDescription.Fields("FABRICANT"))
135         Else
140           itmDescription.SubItems(I_COL_DES_FABRICANT) = vbNullString
145         End If
         
150         itmDescription.SubItems(I_COL_DES_PIECE) = Trim(rstDescription.Fields("PIECE"))
       
155         Call rstDescription.MoveNext
160       Loop
      
165       Call rstDescription.Close
170       Set rstDescription = Nothing
         
175       If lvwDescription.ListItems.count > 0 Then
180         lvwDescription.Visible = True

185         Call lvwDescription.SetFocus
190       Else
195         Call MsgBox("Aucun enregistrement trouvé!")
200       End If
205     End If

210     Exit Sub

AfficherErreur:

215     woups "frmCatalogueMec", "cmdRechercheDescrFR_Click", Err, Erl
End Sub

Private Sub cmdRechercheInventaire_Click()

5       On Error GoTo AfficherErreur

10      Call frmRechercheInventaire.Afficher(MECANIQUE)

15      Exit Sub

AfficherErreur:

20      woups "frmCatalogueMec", "cmdRechercheInventaire_Click", Err, Erl
End Sub

Private Sub cmdRechercheJob_Click()

5       On Error GoTo AfficherErreur

10      Dim rstProjSoum As ADODB.Recordset
15      Dim itmProjSoum As ListItem

20      Screen.MousePointer = vbHourglass

25      Call lvwRechercheJob.ListItems.Clear

30      Set rstProjSoum = New ADODB.Recordset

35      Call rstProjSoum.Open("SELECT DISTINCT IDProjet, SUM(Qté) As QtéTotale FROM GRB_Projet_Pieces WHERE TRIM(NumItem) = '" & Trim$(Replace(txtNoItem.Text, "'", "''")) & "' And Type = 'M' GROUP BY IDProjet", g_connData, adOpenForwardOnly, adLockReadOnly)

40      Do While Not rstProjSoum.EOF
45        Set itmProjSoum = lvwRechercheJob.ListItems.Add

50        itmProjSoum.Text = rstProjSoum.Fields("IDProjet")
55        itmProjSoum.SubItems(I_COL_JOB_QUANTITE) = rstProjSoum.Fields("QtéTotale")

60        Call rstProjSoum.MoveNext
65      Loop

70      Call rstProjSoum.Close

75      Call rstProjSoum.Open("SELECT DISTINCT IDSoumission, SUM(Qté) As QtéTotale FROM GRB_Soumission_Pieces WHERE TRIM(NumItem) = '" & Trim$(Replace(txtNoItem.Text, "'", "''")) & "' And Type = 'M' GROUP BY IDSoumission", g_connData, adOpenForwardOnly, adLockReadOnly)

80      Do While Not rstProjSoum.EOF
85        Set itmProjSoum = lvwRechercheJob.ListItems.Add

90        itmProjSoum.Text = rstProjSoum.Fields("IDSoumission")
95        itmProjSoum.SubItems(I_COL_JOB_QUANTITE) = rstProjSoum.Fields("QtéTotale")

100       Call rstProjSoum.MoveNext
105     Loop

110     Call rstProjSoum.Close
115     Set rstProjSoum = Nothing

120     Screen.MousePointer = vbDefault

125     If lvwRechercheJob.ListItems.count > 0 Then
130       lvwRechercheJob.Visible = True

135       Call lvwRechercheJob.SetFocus
140     Else
145       Call MsgBox("Cette pièce n'a jamais été utilisée dans les jobs!", vbOKOnly)
150     End If

155     Exit Sub

AfficherErreur:

160     woups "frmCatalogueMec", "cmdRechercheJob_Click", Err, Erl
End Sub

Private Sub cmdRechercheAchat_Click()

5       On Error GoTo AfficherErreur

10      Dim rstAchat As ADODB.Recordset
15      Dim itmAchat As ListItem

20      Screen.MousePointer = vbHourglass

25      Call lvwRechercheAchat.ListItems.Clear

30      Set rstAchat = New ADODB.Recordset

35      Call rstAchat.Open("SELECT DISTINCT (IDAchat + '-' + RIGHT('00' & IndexAchat,3)) As NumeroAchat, SUM(Qté) As QtéTotale FROM GRB_Achat_Pieces WHERE TRIM(PIECE) = '" & Trim$(Replace(txtNoItem.Text, "'", "''")) & "' GROUP BY (IDAchat + '-' + RIGHT('00' & IndexAchat,3))", g_connData, adOpenForwardOnly, adLockReadOnly)

40      Do While Not rstAchat.EOF
45        Set itmAchat = lvwRechercheAchat.ListItems.Add

50        itmAchat.Text = rstAchat.Fields("NumeroAchat")
55        itmAchat.SubItems(I_COL_ACHAT_QUANTITE) = rstAchat.Fields("QtéTotale")

60        Call rstAchat.MoveNext
65      Loop

70      Call rstAchat.Close
75      Set rstAchat = Nothing

80      Screen.MousePointer = vbDefault

85      If lvwRechercheAchat.ListItems.count > 0 Then
90       lvwRechercheAchat.Visible = True

95        Call lvwRechercheAchat.SetFocus
100     Else
105       Call MsgBox("Cette pièce n'a jamais été utilisée dans les achats!", vbOKOnly)
110     End If

115     Exit Sub

AfficherErreur:

120     woups "frmCatalogueMec", "cmdRechercheAchat_Click", Err, Erl
End Sub

Private Sub cmdRechercherManufact_Click()

5       On Error GoTo AfficherErreur

10      Dim rstManufact As ADODB.Recordset
15      Dim sManufact   As String
20      Dim itmManufact As ListItem
  
25      sManufact = InputBox("Quel est le manufacturier à rechercher")
  
30      sManufact = Replace(sManufact, "'", "''")
  
35      If sManufact <> vbNullString Then
40        Call lvwFabricant.ListItems.Clear
      
45        Set rstManufact = New ADODB.Recordset
      
50        Call rstManufact.Open("SELECT * FROM GRB_CatalogueMec WHERE INSTR(1, FABRICANT,'" & sManufact & "') > 0 ORDER BY FABRICANT", g_connData, adOpenDynamic, adLockOptimistic)
      
55        Do While Not rstManufact.EOF
60          Set itmManufact = lvwFabricant.ListItems.Add()
        
65          itmManufact.Tag = rstManufact.Fields("CATEGORIE")
        
70          itmManufact.Text = rstManufact.Fields("FABRICANT")
      
75          itmManufact.SubItems(I_COL_MAN_PIECE) = Trim(rstManufact.Fields("PIECE"))
         
80          If Not IsNull(rstManufact.Fields("DESC_FR")) Then
85            itmManufact.SubItems(I_COL_MAN_DESCR_FR) = Trim(rstManufact.Fields("DESC_FR"))
90          Else
95            itmManufact.SubItems(I_COL_MAN_DESCR_FR) = vbNullString
100         End If
      
105         If Not IsNull(rstManufact.Fields("DESC_EN")) Then
110           itmManufact.SubItems(I_COL_MAN_DESCR_EN) = Trim(rstManufact.Fields("DESC_EN"))
115         Else
120           itmManufact.SubItems(I_COL_MAN_DESCR_EN) = vbNullString
125         End If
              
130         Call rstManufact.MoveNext
135       Loop
      
140       Call rstManufact.Close
145       Set rstManufact = Nothing
         
150       If lvwFabricant.ListItems.count > 0 Then
155         lvwFabricant.Visible = True
      
160         Call lvwFabricant.SetFocus
165       Else
170         Call MsgBox("Aucun enregistrement trouvé!")
175       End If
180     End If

185     Exit Sub

AfficherErreur:

190     woups "frmCatalogueMec", "cmdRechercherManufact_Click", Err, Erl
End Sub

Private Sub cmdTotal_Click()

5       On Error GoTo AfficherErreur

10      Dim sAnnee        As String
15      Dim rstTotal      As ADODB.Recordset
20      Dim dblTotalProj  As Double
25      Dim dblTotalAchat As Double

30      sAnnee = InputBox("Pour quelle année? (AAAA)")

35      If Len(sAnnee) = 4 Then
40        If IsNumeric(sAnnee) Then
45          If CInt(sAnnee) <= Year(Date) Then
50            Screen.MousePointer = vbHourglass

55            Set rstTotal = New ADODB.Recordset

60            Call rstTotal.Open("SELECT SUM(Qté) As Total FROM GRB_Projet_Pieces INNER JOIN GRB_ProjetMec ON GRB_Projet_Pieces.IDProjet = GRB_ProjetMec.IDProjet WHERE TRIM(NumItem) = '" & Trim$(Replace(txtNoItem.Text, "'", "''")) & "' AND Type = 'M' AND Left(Creer, 4) = '" & sAnnee & "' AND RIGHT(GRB_Projet_Pieces.IDProjet,2) < '60'", g_connData, adOpenDynamic, adLockOptimistic)

65            If Not IsNull(rstTotal.Fields("Total")) Then
70              dblTotalProj = CDbl(rstTotal.Fields("Total"))
75            Else
80              dblTotalProj = 0
85            End If

90            Call rstTotal.Close

95            Call rstTotal.Open("SELECT SUM(Qté) As Total FROM GRB_Achat_Pieces INNER JOIN GRB_Achat ON GRB_Achat_Pieces.IDAchat = GRB_Achat.IDAchat AND GRB_Achat_Pieces.IndexAchat = GRB_Achat.IndexAchat WHERE TRIM(PIECE) = '" & Trim$(Replace(txtNoItem.Text, "'", "''")) & "' AND Left(DateAchat,4) = '" & sAnnee & "'", g_connData, adOpenDynamic, adLockOptimistic)

100           If Not IsNull(rstTotal.Fields("Total")) Then
105             dblTotalAchat = CDbl(rstTotal.Fields("Total"))
110           Else
115             dblTotalAchat = 0
120           End If

125           Call rstTotal.Close
130           Set rstTotal = Nothing

135           Screen.MousePointer = vbDefault

140           Call MsgBox("Quantité utilisée en " & sAnnee & " : " & vbNewLine & _
                          vbNewLine & _
                          "Projets : " & dblTotalProj & vbNewLine & _
                          "Achats : " & dblTotalAchat)
145         Else
150           Call MsgBox("Année trop grande!", vbOKOnly, "Erreur")
155         End If
160       Else
165         Call MsgBox("Année non numérique!", vbOKOnly, "Erreur")
170       End If
175     Else
180       If Len(sAnnee) <> 0 Then
185         Call MsgBox("L'année doit être sur 4 chiffres!", vbOKOnly, "Erreur")
190       End If
195     End If

200     Exit Sub

AfficherErreur:

205     woups "frmCatalogueMec", "cmdTotal_Click", Err, Erl
End Sub

Private Sub Form_Click()

5       On Error GoTo AfficherErreur

10      lvwDescription.Visible = False
15      lvwFabricant.Visible = False
20      lvwPieces.Visible = False
25      lvwRechercheJob.Visible = False
30      lvwRechercheAchat.Visible = False

35      Exit Sub

AfficherErreur:

40      woups "frmCatalogueMec", "Form_Click", Err, Erl
End Sub

Private Sub frafournisseur_Click()

5       On Error GoTo AfficherErreur

10      lvwDescription.Visible = False
15      lvwFabricant.Visible = False
20      lvwPieces.Visible = False
25      lvwRechercheJob.Visible = False
30      lvwRechercheAchat.Visible = False

35      Exit Sub

AfficherErreur:

40      woups "frmCatalogueMec", "fraFournisseur_Click", Err, Erl
End Sub

Private Sub fraQuantité_Click()

5       On Error GoTo AfficherErreur

10      lvwDescription.Visible = False
15      lvwFabricant.Visible = False
20      lvwPieces.Visible = False
25      lvwRechercheJob.Visible = False
30      lvwRechercheAchat.Visible = False

35      Exit Sub

AfficherErreur:

40      woups "frmCatalogueMec", "fraQuantité_Click", Err, Erl
End Sub



Private Sub lvwCategorie_DblClick()
5
10      Dim itmDescription As ListItem
15      Dim iCompteur      As Integer

20      If lvwCategorie.ListItems.count > 0 Then
25        Screen.MousePointer = vbHourglass

30        Set itmDescription = lvwCategorie.SelectedItem

35        'm_sSelectCategorie = itmDescription.Tag
40        'm_sSelectFabricant = Trim$(itmDescription.SubItems(I_COL_DES_FABRICANT))
45       ' m_sSelectNoItem = Trim$(itmDescription.SubItems(I_COL_DES_PIECE))

50        'If m_eMode = MODE_INACTIF Then
55        '  Call RemplirComboCategorie
60        'Else
65          cmbCategorie.Text = lvwCategorie.SelectedItem.Text
75
            Call cmbCategorie_Click
80        lvwCategorie.Visible = False

85        Screen.MousePointer = vbDefault
90      End If

95      Exit Sub
End Sub

Private Sub lvwCategorie_LostFocus()

5       On Error GoTo AfficherErreur

10      If lvwCategorie.Visible = True Then
15        lvwCategorie.Visible = False
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmCatalogueElec", "lvwCategorie_LostFocus", Err, Erl
End Sub

Private Sub lvwDescription_KeyDown(KeyCode As Integer, Shift As Integer)
        
5       On Error GoTo AfficherErreur

10      If KeyCode = vbKeyReturn Then
15        Call lvwDescription_DblClick
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmCatalogueMec", "lvwDescription_KeyDown", Err, Erl
End Sub

Private Sub lvwFabricant_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

5       On Error GoTo AfficherErreur

10      lvwFabricant.Sorted = True
  
15      If lvwFabricant.SortOrder = lvwAscending Then
20        lvwFabricant.SortOrder = lvwDescending
25      Else
30        lvwFabricant.SortOrder = lvwAscending
35      End If
  
40      lvwFabricant.SortKey = ColumnHeader.Index - 1

45      Exit Sub

AfficherErreur:

50      woups "frmCatalogueMec", "lvwFabricant_ColumnClick", Err, Erl
End Sub

Private Sub lvwDescription_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

5       On Error GoTo AfficherErreur

10      lvwDescription.Sorted = True
  
15      If lvwDescription.SortOrder = lvwAscending Then
20        lvwDescription.SortOrder = lvwDescending
25      Else
30        lvwDescription.SortOrder = lvwAscending
35      End If
  
40      lvwDescription.SortKey = ColumnHeader.Index - 1

45      Exit Sub

AfficherErreur:

50      woups "frmCatalogueMec", "lvwDescription_ColumnClick", Err, Erl
End Sub

Private Sub lvwFabricant_KeyDown(KeyCode As Integer, Shift As Integer)
        
5       On Error GoTo AfficherErreur

10      If KeyCode = vbKeyReturn Then
15        Call lvwFabricant_DblClick
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmCatalogueMec", "lvwFabricant_KeyDown", Err, Erl
End Sub

Private Sub lvwPieces_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

5       On Error GoTo AfficherErreur

10      lvwPieces.Sorted = True
  
15      If lvwPieces.SortOrder = lvwAscending Then
20        lvwPieces.SortOrder = lvwDescending
25      Else
30        lvwPieces.SortOrder = lvwAscending
35      End If
  
40      lvwPieces.SortKey = ColumnHeader.Index - 1

45      Exit Sub

AfficherErreur:

50      woups "frmCatalogueMec", "lvwPieces_ColumnClick", Err, Erl
End Sub

Private Sub lvwPieces_KeyDown(KeyCode As Integer, Shift As Integer)
        
5       On Error GoTo AfficherErreur

10      If KeyCode = vbKeyReturn Then
15        Call lvwPieces_DblClick
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmCatalogueMec", "lvwPieces_KeyDown", Err, Erl
End Sub

Private Sub lvwRechercheJob_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

5       On Error GoTo AfficherErreur

10      If lvwRechercheJob.SortOrder = lvwAscending Then
15        lvwRechercheJob.SortOrder = lvwDescending
20      Else
25        lvwRechercheJob.SortOrder = lvwAscending
30      End If

35      lvwRechercheJob.SortKey = ColumnHeader.Index - 1

40      Exit Sub

AfficherErreur:

45      woups "frmCatalogueMec", "lvwRechercheJob_ColumnClick", Err, Erl
End Sub

Private Sub lvwRechercheAchat_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

5       On Error GoTo AfficherErreur

10      If lvwRechercheAchat.SortOrder = lvwAscending Then
15        lvwRechercheAchat.SortOrder = lvwDescending
20      Else
25        lvwRechercheAchat.SortOrder = lvwAscending
30      End If

35      lvwRechercheAchat.SortKey = ColumnHeader.Index - 1

40      Exit Sub

AfficherErreur:

45      woups "frmCatalogueMec", "lvwRechercheAchat_ColumnClick", Err, Erl
End Sub

Private Sub lvwFabricant_DblClick()

5       On Error GoTo AfficherErreur

10      Dim itmFabricant As ListItem
15      Dim iCompteur    As Integer
  
20      Screen.MousePointer = vbHourglass
  
25      Set itmFabricant = lvwFabricant.SelectedItem

30      m_sSelectCategorie = Trim$(itmFabricant.Tag)
35      m_sSelectFabricant = Trim$(itmFabricant.Text)
40      m_sSelectNoItem = Trim$(itmFabricant.SubItems(I_COL_MAN_PIECE))
    
45      Call RemplirComboCategorie
    
50      lvwFabricant.Visible = False

55      Exit Sub

AfficherErreur:

60      woups "frmCatalogueMec", "lvwFabricant_DblClick", Err, Erl
End Sub

Private Sub lvwDescription_DblClick()

5       On Error GoTo AfficherErreur

10      Dim itmDescription As ListItem
15      Dim iCompteur      As Integer

20      Screen.MousePointer = vbHourglass
  
25      If lvwDescription.ListItems.count > 0 Then
30        Set itmDescription = lvwDescription.SelectedItem

35        If m_eMode = MODE_INACTIF Then
40          m_sSelectCategorie = Trim$(itmDescription.Tag)
45          m_sSelectFabricant = Trim$(itmDescription.SubItems(I_COL_DES_FABRICANT))
50          m_sSelectNoItem = Trim$(itmDescription.SubItems(I_COL_DES_PIECE))

55          Call RemplirComboCategorie
60        Else
65          txtDescriptionFR.Text = itmDescription.Text
70          txtDescriptionEN.Text = itmDescription.SubItems(I_COL_DES_DESCR_EN)
75        End If

80        lvwDescription.Visible = False

85        Screen.MousePointer = vbDefault
90      End If

95      Exit Sub

AfficherErreur:

100     woups "frmCatalogueMec", "lvwDescription_DblClick", Err, Erl
End Sub

Private Sub CmdSupp_Click()

5       On Error GoTo AfficherErreur

10      Dim sDescription As String
15      Dim iCompteur    As Integer
        
        
20      If cmbNoItem.ListCount > 0 Then

25        If MsgBox("Voulez-vous vraiment effacer la pièce " & txtNoItem.Text & "?", vbYesNo) = vbYes Then
30
          

            


            
            If chkInventaire.Value = vbChecked Then
                
35            Call g_connData.Execute("DELETE * FROM GRB_InventaireMec WHERE NoItem = '" & Replace(cmbNoItem.Text, "'", "''") & "'")
40
            End If

            'Efface l'enregistrement de catalogue
            
45          Call g_connData.Execute("DELETE * FROM GRB_CatalogueMec WHERE PIECE = '" & Replace(cmbNoItem.Text, "'", "''") & "'")
           
            'efface l'enr de la table piece frs
50          Call g_connData.Execute("DELETE * FROM GRB_PiecesFRS WHERE PIECE = '" & Replace(cmbNoItem.Text, "'", "''") & "'")
            
55          m_bRempliManuel = True
    
60          m_sSelectNoItem = vbNullString
    
65          If cmbNoItem.ListCount > 1 Then
70            m_sSelectFabricant = cmbFabricant.Text
75          Else
80            m_sSelectFabricant = ""
85          End If
            
90          Call RemplirComboFabricant
            
95          If cmbFabricant.ListCount = 0 Then
100           Call cmbNoItem.Clear
            
105           Call cmbCategorie.RemoveItem(cmbCategorie.ListIndex)
            
110           If cmbCategorie.ListCount > 0 Then
115             cmbCategorie.ListIndex = 0
120           Else
           
125             Call ViderChamps_frs
            
130             Call lvwfournisseur.ListItems.Clear
            
135             Call ViderChamps_piece
           
140           End If
145         End If
      
150         sDescription = txtDescriptionFR.Text
            
155         For iCompteur = 0 To cmbDescriptionFR.ListCount - 1
160           If cmbDescriptionFR.LIST(iCompteur) = sDescription Then
165             m_bBloqueDescription = True

170             cmbDescriptionFR.ListIndex = iCompteur

175             m_bBloqueDescription = False

180             Exit For
           
185           End If
190         Next
195       End If
200     End If

205     Exit Sub

AfficherErreur:

210     woups "frmCatalogueMec", "CmdSupp_Click", Err, Erl
End Sub

Private Sub AfficherItem()

5       On Error GoTo AfficherErreur

        'Affichage de l'enregistrement
10      Dim rstItem       As ADODB.Recordset
15      Dim rstInventaire As ADODB.Recordset
20      Dim iCompteur     As Integer
  
        'Il faut mettre le frame enabled pour vérifier si les CheckBox à l'intérieur
        'sont enabled
25      Call ViderChamps_piece

30      Set rstItem = New ADODB.Recordset

35      Call rstItem.Open("SELECT * FROM GRB_CatalogueMec WHERE PIECE = '" & Replace(m_sNoItem, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
  
        'Si il y a un enregistrement
40      If Not rstItem.EOF Then
          'PIECE_GRB
45        If Not IsNull(rstItem.Fields("PIECE_GRB")) Then
50          txtNoItemGRB.Text = Trim(rstItem.Fields("PIECE_GRB"))
55        Else
60          txtNoItemGRB.Text = vbNullString
65        End If

          'DESCR_EN
70        If Not IsNull(rstItem.Fields("DESC_EN").Value) Then
71          txtDescriptionEN.Text = Trim(rstItem.Fields("DESC_EN"))
72        Else
73          txtDescriptionEN.Text = vbNullString
74        End If

          'FABRICANT
80        If Not IsNull(rstItem.Fields("FABRICANT").Value) Then
81          txtFabricant.Text = Trim(rstItem.Fields("FABRICANT"))
82        Else
83          txtFabricant.Text = vbNullString
84        End If

          'DESCR_FR
95        If Not IsNull(rstItem.Fields("DESC_FR")) Then
100         For iCompteur = 0 To cmbDescriptionFR.ListCount - 1
105           If cmbDescriptionFR.LIST(iCompteur) = rstItem.Fields("DESC_FR") Then
110             cmbDescriptionFR.ListIndex = iCompteur
          
115             Exit For
120           End If
125         Next
130       Else
135         If cmbDescriptionFR.ListIndex = -1 Then
140           Call cmbDescriptionFR_Click
145         Else
150           cmbDescriptionFR.ListIndex = -1
155         End If
160       End If
  
          'COMMENT
165       If Not IsNull(rstItem.Fields("COMMENTAIRE")) Then
170         txtComment.Text = Trim(rstItem.Fields("COMMENTAIRE"))
175       Else
180         txtComment.Text = vbNullString
185       End If
   
190       Call RemplirListViewFournisseur
195     Else
200       Call MsgBox("Impossible de trouver la pièce!", vbOKOnly, "Erreur")
205     End If

210     Call rstItem.Close
215     Set rstItem = Nothing

220     Set rstInventaire = New ADODB.Recordset

225     Call rstInventaire.Open("SELECT * FROM GRB_InventaireMec WHERE NoItem = '" & Replace(txtNoItem.Text, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)

230     If Not rstInventaire.EOF Then
235       chkInventaire.Value = vbChecked

240       chkBoite.Value = Abs(CInt(rstInventaire.Fields("CommandeParBoite")))

245       If chkBoite.Value = vbChecked Then
250         txtQuantitéBoite.Text = rstInventaire.Fields("QteBoite")
255       End If

260       For iCompteur = 0 To cmbLocalisation.ListCount - 1
265         If cmbLocalisation.LIST(iCompteur) = rstInventaire.Fields("Localisation") Then
270           cmbLocalisation.ListIndex = iCompteur

275           Exit For
280         End If
285       Next

290       txtQuantiteStock.Text = rstInventaire.Fields("QuantitéStock")
295       chkMinimum.Value = Abs(CInt(rstInventaire.Fields("Minimum")))
300       txtQuantiteMinimum.Text = rstInventaire.Fields("QuantitéMinimum")
305       txtQuantiteCommande.Text = rstInventaire.Fields("Commande")
310     End If

315     Call rstInventaire.Close
320     Set rstInventaire = Nothing

325     Exit Sub

AfficherErreur:

330     woups "frmCatalogueMec", "AfficherItem", Err, Erl
End Sub

Private Sub AfficherFRS()

5       On Error GoTo AfficherErreur

        'Affichage de l'enregistrement
10      Dim rstItemFRS As ADODB.Recordset
15      Dim iCompteur  As Integer
  
20      Set rstItemFRS = New ADODB.Recordset
  
25      Call rstItemFRS.Open("SELECT * FROM GRB_PiecesFRS WHERE NoEnreg = " & lvwfournisseur.SelectedItem.Tag, g_connData, adOpenDynamic, adLockOptimistic)
  
        'Si le champs est Enabled.. c'est parce que le champs existe dans la table
        
30      For iCompteur = 0 To cmbfrs.ListCount - 1
35        If cmbfrs.LIST(iCompteur) = lvwfournisseur.SelectedItem.Text Then
40          cmbfrs.ListIndex = iCompteur
            
45          Exit For
50        End If
55      Next

        'PERS_RESS
60      If Not IsNull(rstItemFRS.Fields("PERS_RESS")) Then
65        For iCompteur = 0 To cmbPersRess.ListCount - 1
70          If cmbPersRess.ItemData(iCompteur) = rstItemFRS.Fields("PERS_RESS") Then
75             cmbPersRess.ListIndex = iCompteur
        
80            Exit For
85          End If
90        Next
95      Else
100       cmbPersRess.ListIndex = -1
105     End If

        'PRIX_LIST
110     If Not IsNull(rstItemFRS.Fields("PRIX_LIST")) Then
115       txtPrixList.Text = rstItemFRS.Fields("PRIX_LIST")
120     Else
125       txtPrixList.Text = vbNullString
130     End If
  
        'ESCOMPTE
135     If Not IsNull(rstItemFRS.Fields("ESCOMPTE")) Then
140       mskEscompte.Text = rstItemFRS.Fields("ESCOMPTE")
145     Else
150       mskEscompte.Text = vbNullString
155     End If

        'PRIX_NET
160     If Not IsNull(rstItemFRS.Fields("PRIX_NET")) Then
165       txtPrixNet.Text = rstItemFRS.Fields("PRIX_NET")
170     Else
175       txtPrixNet.Text = vbNullString
180     End If
    
        'PRIX_SP
185     If Not IsNull(rstItemFRS.Fields("PRIX_SP")) Then
190       txtPrixSpecial.Text = rstItemFRS.Fields("PRIX_SP")
195     Else
200       txtPrixSpecial.Text = vbNullString
205     End If
    
        'VALIDE
210     If Not IsNull(rstItemFRS.Fields("VALIDE")) Then
215       mskValide.Text = rstItemFRS.Fields("VALIDE")
220     Else
225       mskValide.Text = vbNullString
230     End If
  
        'QUOTER
235     If rstItemFRS.Fields("quoter") = True Then
240       chkquoter.Value = vbChecked
245     Else
250       chkquoter.Value = vbUnchecked
255     End If
  
        'devise monetaire
260     If rstItemFRS.Fields("DeviseMonétaire") = "CAN" Then
265       optCAN.Value = True
270     Else
275       If rstItemFRS.Fields("DeviseMonétaire") = "USA" Then
280         optUSA.Value = True
285       Else
290         optSpain.Value = True
295       End If
300     End If
  
        'affiche drapeau
305     Call AfficherDrapeau
  
310     Call rstItemFRS.Close
315     Set rstItemFRS = Nothing

320     Exit Sub

AfficherErreur:

325     woups "frmCatalogueMec", "AfficherFRS", Err, Erl
End Sub

Private Sub cmbNoItem_Click()

5       On Error GoTo AfficherErreur

        'Affichage de l'enregistrement
10      Screen.MousePointer = vbHourglass
  
        'Il faut mettre le nom de l'élément sélectionné dans le textbox pour ensuite
        'l'utiliser pour les requêtes SQL
15      txtNoItem.Text = cmbNoItem.Text
  
20      m_sNoItem = txtNoItem.Text
  
25      m_bBloqueDescription = True
  
30      Call AfficherItem
  
35      m_bBloqueDescription = False
  
        'remplir combo frs
40      Call RemplirComboFRS
  
45      Screen.MousePointer = vbDefault

50      Exit Sub

AfficherErreur:

55      woups "frmCatalogueMec", "cmbNoItem_Click", Err, Erl
End Sub

Private Sub cmbFabricant_Click()

5      On Error GoTo AfficherErreur

        'quand un manufact est selectionné on remplir le combo des NumItem
10      Screen.MousePointer = vbHourglass
  
15      txtFabricant.Text = cmbFabricant.Text

20      Call RemplirComboDescription

25      m_bBloqueDescription = True
  
30      If m_bRempliManuel = True Then
35
        Call RemplirComboNoItem
        
40        m_bRempliManuel = False
45      Else
        
50        Call RemplirComboNoItem
        
55      End If
  
60      m_bBloqueDescription = False
  
65      Screen.MousePointer = vbDefault
        If sChoisirTous = ")" Then
        Call RemplirComboCategorie
        End If
        
        
70      Exit Sub

AfficherErreur:

75      woups "frmCatalogueMec", "cmbFabricant_Click", Err, Erl
End Sub

Private Sub cmdSuppFrs_Click()

5       On Error GoTo AfficherErreur

10      If lvwfournisseur.ListItems.count > 0 Then
15        Call SupprimerFournisseur
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmCatalogueMec", "cmdSuppFrs_Click", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

10      Call Unload(frmChoixCatalogue)

15      Call ActiverBoutonsGroupe
  
20      m_bBloqueDescription = True
  
25      Set m_collPieceDescFR = New Collection
  
        'Barrer les champs
30      Call BarrerChamps_piece(True)
    
        'Activer ou désactiver certains controles
35      Call MontrerControles(MODE_INACTIF)
        
40      Call RemplirComboLocalisation
        
        'Remplir le combo des pièces disponibles
45      Call RemplirComboCategorie
    
50      m_bBloqueDescription = False

55      Exit Sub

AfficherErreur:

60      woups "frmCatalogueMec", "Form_Load", Err, Erl
End Sub

Private Sub ActiverBoutonsGroupe()

5       On Error GoTo AfficherErreur

10      CmdAdd.Enabled = g_bModificationCatalogueMec
15      CmdSupp.Enabled = g_bModificationCatalogueMec
20      CmdModif.Enabled = g_bModificationCatalogueMec
25      cmdAddFrs.Enabled = g_bModificationCatalogueMec
30      cmdSuppFrs.Enabled = g_bModificationCatalogueMec
35      cmdModifFrs.Enabled = g_bModificationCatalogueMec
40      cmdDemande.Enabled = g_bModificationCatalogueMec
          
45      Exit Sub

AfficherErreur:

50      woups "frmCatalogueMec", "ActiverBoutonsGroupe", Err, Erl
End Sub

Public Sub RemplirComboFabricant()

5       On Error GoTo AfficherErreur

        'Rempli le combo des fabricants
10      Dim rstFabricant As ADODB.Recordset
15      Dim sCategorie   As String
20      Dim iCompteur    As Integer
  
25      sCategorie = Replace(cmbCategorie.Text, "'", "''")

30      Set rstFabricant = New ADODB.Recordset
        
35      Call rstFabricant.Open("SELECT DISTINCT FABRICANT FROM GRB_CatalogueMec WHERE CATEGORIE = '" & sCategorie & "' ORDER BY FABRICANT", g_connData, adOpenDynamic, adLockOptimistic)
        
        'Il faut vider le combo avant de le remplir
40      Call cmbFabricant.Clear
42      sChoisirTous = ""
43      'on ajoute la possibilité de choisir tout les fabricants
44      Call cmbFabricant.AddItem("-- CHOISIR TOUS --")
        If Not rstFabricant.EOF Then rstFabricant.MoveFirst
        
        'Tant que ce n'est pas la fin des enregistrements
46      Do While Not rstFabricant.EOF
          'Si l'élément n'est pas null
50        If Not IsNull(rstFabricant.Fields("Fabricant")) Then
            'on l'ajoute
            
55          Call cmbFabricant.AddItem(Trim(rstFabricant.Fields("FABRICANT")))
            
            If sChoisirTous = "" Then
                sChoisirTous = " AND (FABRICANT = '" & Trim(rstFabricant.Fields("FABRICANT")) & "'"
            Else
56              sChoisirTous = sChoisirTous + " OR FABRICANT = '" & Trim(rstFabricant.Fields("FABRICANT")) & "'"
            End If
60        End If
  
65        Call rstFabricant.MoveNext
70      Loop
        
        sChoisirTous = sChoisirTous + ")"
        

75      Call rstFabricant.Close
80      Set rstFabricant = Nothing
        
        'Si le combo n'est pas vide, on sélectionne le premier élément
      
85      If cmbFabricant.ListCount > 0 Then
            
90        If m_sSelectFabricant <> vbNullString Then
             
95          For iCompteur = 0 To cmbFabricant.ListCount - 1
100           If UCase(cmbFabricant.LIST(iCompteur)) = UCase(m_sSelectFabricant) Then
105             cmbFabricant.ListIndex = iCompteur

110             m_sSelectFabricant = ""
              
115             Exit For
       
120           End If
125         Next
130       Else
            
135         cmbFabricant.ListIndex = 0
            
140       End If
145     Else
        
150       Call cmbNoItem.Clear
155       Call cmbDescriptionFR.Clear
160     End If
        

165     Exit Sub

AfficherErreur:

170     woups "frmCatalogueMec", "RemplirComboFabricant", Err, Erl
End Sub

Private Sub cmdRechercherPiece_Click()

5       On Error GoTo AfficherErreur

10      Dim rstPiece    As ADODB.Recordset
15      Dim sPiece      As String
20      Dim itmPiece    As ListItem
25      Dim iCompteur   As Integer
30      Dim sPieceModif As String
35      Dim sLettre     As String
    
40      sPiece = InputBox("Quelle est la pièce à rechercher?")
    
45      If sPiece <> vbNullString Then
50        For iCompteur = 1 To Len(sPiece)
55          sLettre = Mid$(sPiece, iCompteur, 1)
        
60          If (Asc(sLettre) >= 48 And Asc(sLettre) <= 57) Or _
               (Asc(sLettre) >= 65 And Asc(sLettre) <= 90) Or _
               (Asc(sLettre) >= 97 And Asc(sLettre) <= 122) Then
65            sPieceModif = sPieceModif & sLettre
70          End If
75        Next

80        Call lvwPieces.ListItems.Clear
    
85        Set rstPiece = New ADODB.Recordset
    
90        Call rstPiece.Open("SELECT * FROM GRB_CatalogueMec WHERE INSTR(1, PIECE_MODIF, '" & sPieceModif & "') > 0 ORDER BY PIECE", g_connData, adOpenDynamic, adLockOptimistic)
      
95        Do While Not rstPiece.EOF
100         Set itmPiece = lvwPieces.ListItems.Add
        
105         itmPiece.Text = rstPiece.Fields("PIECE")

110         If Not IsNull(rstPiece.Fields("FABRICANT")) Then
115           itmPiece.SubItems(I_COL_PIECE_FABRICANT) = rstPiece.Fields("FABRICANT")
120         Else
125           itmPiece.SubItems(I_COL_PIECE_FABRICANT) = vbNullString
130         End If
            
135         If Not IsNull(rstPiece.Fields("DESC_FR")) Then
140           itmPiece.SubItems(I_COL_PIECE_DESCR_FR) = rstPiece.Fields("DESC_FR")
145         Else
150           itmPiece.SubItems(I_COL_PIECE_DESCR_FR) = vbNullString
155         End If
     
160         If Not IsNull(rstPiece.Fields("DESC_EN")) Then
165           itmPiece.SubItems(I_COL_PIECE_DESCR_EN) = rstPiece.Fields("DESC_EN")
170         Else
175           itmPiece.SubItems(I_COL_PIECE_DESCR_EN) = vbNullString
180         End If
       
185         itmPiece.Tag = rstPiece.Fields("CATEGORIE")
        
190         Call rstPiece.MoveNext
195       Loop
      
200       Call rstPiece.Close
205       Set rstPiece = Nothing
        
210       If lvwPieces.ListItems.count > 0 Then
215         lvwPieces.Visible = True
      
220         Call lvwPieces.SetFocus
225       Else
230         Call MsgBox("Aucun enregistrement trouvé!")
235       End If
240     End If

245     Exit Sub

AfficherErreur:

250     woups "frmCatalogueMec", "cmdRechercherPiece_Click", Err, Erl
End Sub

Public Sub AfficherForm(ByVal sCategorie As String, ByVal sNomFab As String, ByVal sNoItem As String)

5       On Error GoTo AfficherErreur
        
        'Ouverture de la fenêtre
10      Dim iCompteur As Integer
  
        'Barrer les champs
15      Call BarrerChamps_piece(True)
    
        'Activer ou désactiver certains controles
20      Call MontrerControles(MODE_INACTIF)
  
        'Remplir le combo des pièces disponibles
25      Call RemplirComboCategorie
  
  
30      If sCategorie <> "" Then
35        For iCompteur = 0 To cmbCategorie.ListCount - 1
40          If cmbCategorie.LIST(iCompteur) = sCategorie Then
45            cmbCategorie.ListIndex = iCompteur

50            Exit For
55          End If
60        Next
65      End If
  
70      If sNomFab <> "" Then
75        For iCompteur = 0 To cmbFabricant.ListCount - 1
80          If cmbFabricant.LIST(iCompteur) = sNomFab Then
85            cmbFabricant.ListIndex = iCompteur

90            Exit For
95          End If
100       Next
105     End If
  
110     If sNoItem <> "" Then
115       For iCompteur = 0 To cmbNoItem.ListCount - 1
120         If cmbNoItem.LIST(iCompteur) = sNoItem Then
125           cmbNoItem.ListIndex = iCompteur

130           Exit For
135         End If
140       Next
145     End If
  
150     Call Me.Show

155     Exit Sub

AfficherErreur:

160     woups "frmCatalogueMec", "AfficherForm", Err, Erl
End Sub

Public Sub RemplirComboNoItem()

5       On Error GoTo AfficherErreur

        'Rempli le combo de numéros d'item
10      Dim rstNoItem  As ADODB.Recordset
15      Dim sCategorie As String
20      Dim iCompteur  As Integer
25      Dim sFabricant As String
  
30      sCategorie = Replace(cmbCategorie.Text, "'", "''")
35      sFabricant = Replace(cmbFabricant.Text, "'", "''")
  
40      Set rstNoItem = New ADODB.Recordset

41      If sFabricant = "-- CHOISIR TOUS --" Then
            If cmbCategorie.Text = "Équipements et outillages - Sécurité/nettoyage" Or cmbCategorie.Text = "Quincaillerie - Boulons Hex. (Bolts) 1/4-20" Or cmbCategorie.Text = "Quincaillerie - Rondelle plate (Washer)" Or sChoisirTous = ")" Then
                Call rstNoItem.Open("SELECT PIECE FROM GRB_CatalogueMec WHERE CATEGORIE = '" & sCategorie & "' ORDER BY TRIM(PIECE)", g_connData, adOpenDynamic, adLockOptimistic)
            Else
54              Call rstNoItem.Open("SELECT PIECE FROM GRB_CatalogueMec WHERE CATEGORIE = '" & sCategorie & "'" & sChoisirTous & " ORDER BY TRIM(PIECE)", g_connData, adOpenDynamic, adLockOptimistic)
55          End If


        Else
56          Call rstNoItem.Open("SELECT PIECE FROM GRB_CatalogueMec WHERE CATEGORIE = '" & sCategorie & "' AND FABRICANT = '" & sFabricant & "' ORDER BY TRIM(PIECE)", g_connData, adOpenDynamic, adLockOptimistic)
57      End If


        'Il faut vider le combo avant de le remplir
58      Call cmbNoItem.Clear

        'Tant que c'est n'est pas la fin des enregistrements
59      Do While Not rstNoItem.EOF
          'Si le champs n'est pas vide
60        If Not IsNull(rstNoItem.Fields("PIECE")) Then
            'On l'ajoute
65          Call cmbNoItem.AddItem(Trim(rstNoItem.Fields("PIECE")))
70        End If
  
75        Call rstNoItem.MoveNext
80      Loop

85      Call rstNoItem.Close
90      Set rstNoItem = Nothing

        'Si le combo n'est pas vide, on sélectionne le premier élément
95      If cmbNoItem.ListCount > 0 Then
100       If m_sSelectNoItem <> vbNullString Then
105         For iCompteur = 0 To cmbNoItem.ListCount - 1
110           If cmbNoItem.LIST(iCompteur) = m_sSelectNoItem Then
115             cmbNoItem.ListIndex = iCompteur

120             m_sSelectNoItem = ""
                
125             Exit For
130           End If
135         Next
140       Else
145         cmbNoItem.ListIndex = 0
150       End If
155     End If

160     Exit Sub

AfficherErreur:

165     woups "frmCatalogueMec", "RemplirComboNoItem", Err, Erl
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
  
75      Call rstPieceFRS.Open("SELECT PrixReel, PRIX_NET, PRIX_SP, DeviseMonétaire FROM GRB_PiecesFRS WHERE PIECE = '" & Replace(sNoItem, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)
  
80      Do While Not rstPieceFRS.EOF
85        If rstPieceFRS.Fields("PRIX_NET") <> vbNullString Then
90          sPrixCalcul = Replace(rstPieceFRS.Fields("PRIX_NET"), ".", ",")
95        Else
100         sPrixCalcul = Replace(rstPieceFRS.Fields("PRIX_SP"), ".", ",")
105       End If
    
110       If rstPieceFRS.Fields("DeviseMonétaire") = "USA" Then
115         rstPieceFRS.Fields("PrixReel") = Conversion(CStr(Round(CDbl(sPrixCalcul) / CDbl(sTauxUSA), 4)), MODE_DECIMAL, 4)
120       Else
125         If rstPieceFRS.Fields("DeviseMonétaire") = "SPA" Then
130           rstPieceFRS.Fields("PrixReel") = Conversion(CStr(Round(CDbl(sPrixCalcul) / CDbl(sTauxSPA), 4)), MODE_DECIMAL, 4)
135         Else
140           rstPieceFRS.Fields("PrixReel") = Conversion(sPrixCalcul, MODE_DECIMAL, 4)
145         End If
150       End If
   
155       Call rstPieceFRS.Update
  
160       Call rstPieceFRS.MoveNext
165     Loop
  
170     Call rstPieceFRS.Close
175     Set rstPieceFRS = Nothing

180     Exit Sub

AfficherErreur:

185     woups "frmCatalogueMec", "CalculerPrixReel", Err, Erl
End Sub

Public Sub RemplirListViewFournisseur()

5       On Error GoTo AfficherErreur
        
        ''''''''''''''''''''''''''''''
        ' Remplis lister fournisseur '
        ''''''''''''''''''''''''''''''
10      Dim rstPieceFRS As ADODB.Recordset
15      Dim rstContact  As ADODB.Recordset
20      Dim iCompteur   As Integer
25      Dim itmFRS      As ListItem
30      Dim lCouleur    As Long
  
        'vide le lister
35      Call lvwfournisseur.ListItems.Clear
  
40      Call CalculerPrixReel(txtNoItem.Text)
  
45      Set rstPieceFRS = New ADODB.Recordset
  
50      Call rstPieceFRS.Open("SELECT GRB_PiecesFRS.*, GRB_Fournisseur.NomFournisseur FROM GRB_PiecesFRS INNER JOIN GRB_Fournisseur ON GRB_PiecesFRS.IDFRS = GRB_Fournisseur.IDFRS WHERE PIECE = '" & Replace(cmbNoItem.Text, "'", "''") & "' AND Type = 'M' ORDER BY PrixReel", g_connData, adOpenDynamic, adLockOptimistic)
  
55      Set rstContact = New ADODB.Recordset
  
        'Tant qu'il y a des fournisseurs de la pièce, ajoute dans ListView
60      Do While Not rstPieceFRS.EOF
65        If rstPieceFRS.Fields("DeviseMonétaire") = "CAN" Then
70          lCouleur = COLOR_ROUGE
75        Else
80          lCouleur = COLOR_BLEU
85        End If
        
90        Set itmFRS = lvwfournisseur.ListItems.Add
          
95        itmFRS.Text = rstPieceFRS.Fields("NomFournisseur")
100       itmFRS.ForeColor = lCouleur
      
105       itmFRS.Tag = rstPieceFRS.Fields("NoEnreg")
                      
110       If Trim(rstPieceFRS.Fields("PERS_RESS")) <> vbNullString Then
115         Call rstContact.Open("SELECT NomContact FROM GRB_Contact WHERE IDContact = " & rstPieceFRS.Fields("PERS_RESS"), g_connData, adOpenDynamic, adLockOptimistic)
            
120         If Not rstContact.EOF Then
125           itmFRS.SubItems(I_COL_FRS_PERS_RESS) = rstContact.Fields("NomContact")
130         Else
135           itmFRS.SubItems(I_COL_FRS_PERS_RESS) = vbNullString
140         End If

145         itmFRS.ListSubItems(I_COL_FRS_PERS_RESS).ForeColor = lCouleur
        
150         Call rstContact.Close
155       End If
               
160       If Not IsNull(rstPieceFRS.Fields("Date")) Then
165         itmFRS.SubItems(I_COL_FRS_DATE) = rstPieceFRS.Fields("Date")
170       Else
175         itmFRS.SubItems(I_COL_FRS_DATE) = vbNullString
180       End If
                      
185       itmFRS.ListSubItems(I_COL_FRS_DATE).ForeColor = lCouleur
                            
190       If Not IsNull(rstPieceFRS.Fields("Entrer_par")) Then
195         itmFRS.SubItems(I_COL_FRS_ENTRER_PAR) = rstPieceFRS.Fields("entrer_par")
200       Else
205         itmFRS.SubItems(I_COL_FRS_ENTRER_PAR) = vbNullString
210       End If
    
215       itmFRS.ListSubItems(I_COL_FRS_ENTRER_PAR).ForeColor = lCouleur
                 
220       If Not IsNull(rstPieceFRS.Fields("Valide")) Then
225         itmFRS.SubItems(I_COL_FRS_VALIDE) = rstPieceFRS.Fields("Valide")
230         itmFRS.ListSubItems(I_COL_FRS_VALIDE).ForeColor = lCouleur
235       Else
240         itmFRS.SubItems(I_COL_FRS_VALIDE) = vbNullString
245       End If
             
250       If Not IsNull(rstPieceFRS.Fields("PRIX_LIST")) Or rstPieceFRS.Fields("PRIX_LIST") <> vbNullString Then
255         itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = Conversion(rstPieceFRS.Fields("prix_list"), MODE_ARGENT, 4)

260         itmFRS.ListSubItems(I_COL_FRS_PRIX_LIST).ForeColor = lCouleur
265       Else
270         itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = vbNullString
275       End If
    
280       If Trim(rstPieceFRS.Fields("ESCOMPTE")) <> vbNullString Then
            'Enlève les "_", met un format pourcentage et remplace les "." par des ","
285         itmFRS.SubItems(I_COL_FRS_ESCOMPTE) = Conversion(Replace(rstPieceFRS.Fields("ESCOMPTE"), "_", vbNullString), MODE_POURCENT)
290       Else
295         itmFRS.SubItems(I_COL_FRS_ESCOMPTE) = vbNullString
300       End If
     
305       itmFRS.ListSubItems(I_COL_FRS_ESCOMPTE).ForeColor = lCouleur
    
310       If rstPieceFRS.Fields("PRIX_NET") <> vbNullString Then
315         itmFRS.SubItems(I_COL_FRS_PRIX_NET) = Conversion(Round(Replace(rstPieceFRS.Fields("PRIX_NET"), ".", ","), 4), MODE_ARGENT, 4)
320       Else
325         itmFRS.SubItems(I_COL_FRS_PRIX_NET) = vbNullString
330       End If
            
335       itmFRS.ListSubItems(I_COL_FRS_PRIX_NET).ForeColor = lCouleur
              
340       If rstPieceFRS.Fields("PRIX_SP") <> vbNullString Then
345         itmFRS.SubItems(I_COL_FRS_PRIX_SP) = Conversion(Round(rstPieceFRS.Fields("PRIX_SP"), 4), MODE_ARGENT, 4)
350       Else
355         itmFRS.SubItems(I_COL_FRS_PRIX_SP) = vbNullString
360       End If
    
365       itmFRS.ListSubItems(I_COL_FRS_PRIX_SP).ForeColor = lCouleur
      
370       If rstPieceFRS.Fields("QUOTER") = True Then
375         itmFRS.SubItems(I_COL_FRS_QUOTER) = "Oui"
380       Else
385         itmFRS.SubItems(I_COL_FRS_QUOTER) = "Non"
390       End If
      
395       itmFRS.ListSubItems(I_COL_FRS_QUOTER).ForeColor = lCouleur
     
400       Call rstPieceFRS.MoveNext
405     Loop
      
410     Call lvwfournisseur.Refresh
  
        'Ferme la table
415     Call rstPieceFRS.Close
420     Set rstPieceFRS = Nothing

425     Set rstContact = Nothing

430     Exit Sub

AfficherErreur:

435     woups "frmCatalogueMec", "RemplirListViewFournisseur", Err, Erl
End Sub

Private Sub lvwDescription_LostFocus()

5       On Error GoTo AfficherErreur

10      If lvwDescription.Visible = True Then
15        lvwDescription.Visible = False
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmCatalogueMec", "lvwDescription_LostFocus", Err, Erl
End Sub

Private Sub lvwRechercheJob_LostFocus()

5       On Error GoTo AfficherErreur

10      If lvwRechercheJob.Visible = True Then
15        lvwRechercheJob.Visible = False
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmCatalogueMec", "lvwRechercheJob_LostFocus", Err, Erl
End Sub

Private Sub lvwRechercheAchat_LostFocus()

5       On Error GoTo AfficherErreur

10      If lvwRechercheAchat.Visible = True Then
15        lvwRechercheAchat.Visible = False
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmCatalogueMec", "lvwRechercheAchat_LostFocus", Err, Erl
End Sub

Private Sub lvwFabricant_LostFocus()

5       On Error GoTo AfficherErreur

10      If lvwFabricant.Visible = True Then
15        lvwFabricant.Visible = False
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmCatalogueMec", "lvwFabricant_LostFocus", Err, Erl
End Sub

Private Sub lvwFournisseur_DblClick()

5       On Error GoTo AfficherErreur

        'modifie un fournisseur pour la piece
10      If lvwfournisseur.ListItems.count > 0 Then
15        Call ModifierFournisseur
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmCatalogueMec", "lvwFournisseur_DblClick", Err, Erl
End Sub

Private Sub lvwfournisseur_KeyDown(KeyCode As Integer, Shift As Integer)

5       On Error GoTo AfficherErreur

10      If lvwfournisseur.ListItems.count > 0 Then
15        If KeyCode = vbKeyDelete Then
20          Call SupprimerFournisseur
25        End If
30      End If

35      Exit Sub

AfficherErreur:

40      woups "frmCatalogueMec", "lvwfournisseur_KeyDown", Err, Erl
End Sub

Private Sub ModifierFournisseur()

5       On Error GoTo AfficherErreur

10      Call BarrerChamps_piece(True)
  
        'affiche pour entre des valeurs
15      Call MontrerControles(MODE_AJOUT_MODIF_FRS)

20      m_bAjout = False

        'affiche les données du fournisseur sélectionné
25      Call AfficherFRS

30      Exit Sub

AfficherErreur:

35      woups "frmCatalogueMec", "ModifierFournisseur", Err, Erl
End Sub

Private Sub SupprimerFournisseur()

5       On Error GoTo AfficherErreur

10      If MsgBox("Voulez-vous vraiment effacer le fournisseur " & lvwfournisseur.SelectedItem.Text & "?", vbYesNo) = vbYes Then
          'Efface l'enregistrement de la table GRB_PiecesFRS
15        Call g_connData.Execute("DELETE * FROM GRB_PiecesFRS WHERE NoEnreg = " & lvwfournisseur.SelectedItem.Tag)
      
          'Remplir le ListView des fournisseurs
20        Call RemplirListViewFournisseur
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmCatalogueMec", "SupprimerFournisseur", Err, Erl
End Sub

Private Sub lvwPieces_DblClick()

5      On Error GoTo AfficherErreur

10      Dim itmPieces As ListItem
15      Dim iCompteur As Integer
  
20      Screen.MousePointer = vbHourglass
  
25      Set itmPieces = lvwPieces.SelectedItem

30      m_sSelectCategorie = Trim$(itmPieces.Tag)
35      m_sSelectFabricant = Trim$(itmPieces.SubItems(I_COL_PIECE_FABRICANT))
40      m_sSelectNoItem = Trim$(itmPieces.Text)
  
45      Call RemplirComboCategorie
    
50      lvwPieces.Visible = False

55      Screen.MousePointer = vbDefault

60      Exit Sub

AfficherErreur:

65      woups "frmCatalogueMec", "lvwPieces_DblClick", Err, Erl
End Sub

Private Sub lvwPieces_LostFocus()

5       On Error GoTo AfficherErreur

10      If lvwPieces.Visible = True Then
15        lvwPieces.Visible = False
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmCatalogueMec", "lvwPieces_LostFocus", Err, Erl
End Sub

Private Sub mskEscompte_GotFocus()

5       On Error GoTo AfficherErreur

        'Quand le maskEdit prend le focus, on set le masque
10      If mskEscompte.Enabled = True Then
15        mskEscompte.mask = "0,####"
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmCatalogueMec", "mskEscompte_GotFocus", Err, Erl
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

40      woups "frmCatalogueMec", "mskEscompte_LostFocus", Err, Erl
End Sub

Private Sub optCAN_Click()

5       On Error GoTo AfficherErreur
              'dependant la devise, affiche le drapeau
10      Call AfficherDrapeau

15      Exit Sub

AfficherErreur:

20      woups "frmCatalogueMec", "optCAN_Click", Err, Erl
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

30        lblDevise1.Visible = False
35        txtTauxChange.Visible = False
40        lblDevise2.Visible = False
45      Else
50        If optUSA.Value = True Then
55          imgEU.Visible = True
60          imgCanada.Visible = False
65          imgSpain.Visible = False
70        Else
75          imgSpain.Visible = True
80          imgCanada.Visible = False
85          imgEU.Visible = False
90        End If

95        Call AfficherTauxChange
100     End If

105     Exit Sub

AfficherErreur:

110     woups "frmCatalogueMec", "AfficherDrapeau", Err, Erl
End Sub

Private Sub AfficherTauxChange()

5       On Error GoTo AfficherErreur

10      Dim rstConfig As ADODB.Recordset

15      Set rstConfig = New ADODB.Recordset

20      Call rstConfig.Open("SELECT TauxAmericain, TauxEspagnol FROM GRB_Config", g_connData, adOpenDynamic, adLockOptimistic)

25      If optUSA.Value = True Then
30        lblDevise2.Caption = "$ USA"
35        txtTauxChange.Text = rstConfig.Fields("TauxAmericain")
40      Else
45        lblDevise2.Caption = "$ SPA"
50        txtTauxChange.Text = rstConfig.Fields("TauxEspagnol")
55      End If

60      lblDevise1.Visible = True
65      txtTauxChange.Visible = True
70      lblDevise2.Visible = True

75      Call rstConfig.Close
80      Set rstConfig = Nothing

85      Exit Sub

AfficherErreur:

90      woups "frmCatalogueMec", "AfficherTauxChange", Err, Erl
End Sub

Private Sub optSpain_Click()

5       On Error GoTo AfficherErreur

        'dependant la devise, affiche le drapeau
10      Call AfficherDrapeau

15      Exit Sub

AfficherErreur:

20      woups "frmCatalogueMec", "optSpain_Click", Err, Erl
End Sub

Private Sub optUSA_Click()

5       On Error GoTo AfficherErreur

        'dependant la devise, affiche le drapeau
10      Call AfficherDrapeau

15      Exit Sub

AfficherErreur:

20      woups "frmCatalogueMec", "optUSA_Click", Err, Erl
End Sub

Private Sub txtNoItem_Change()

5       On Error GoTo AfficherErreur

10      If m_eMode = MODE_AJOUT_MODIF_ELEC Then
15        txtNoItemGRB.Text = txtNoItem.Text & "GRB"
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmCatalogueMec", "txtNoItem_Change", Err, Erl
End Sub

Private Sub txtPrixList_LostFocus()

5       On Error GoTo AfficherErreur

10      If txtPrixList.Text <> vbNullString Then
15        txtPrixList.Text = Replace(txtPrixList, ".", ",")
  
20        If Not IsNumeric(txtPrixList.Text) Then
25          Call MsgBox("Valeur non numérique!", vbOKOnly, "Erreur")
30          txtPrixList.Text = vbNullString
35        End If
40      End If

45      Exit Sub

AfficherErreur:

50      woups "frmCatalogueMec", "txtPrixList_LostFocus", Err, Erl
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

45      woups "frmCatalogueMec", "txtPrixNet_Change", Err, Erl
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

60      woups "frmCatalogueMec", "txtPrixSpecial_Change", Err, Erl
End Sub

Private Sub txtPrixNet_GotFocus()

5       On Error GoTo AfficherErreur

        'Si le prix net prend le focus
10      Call CalculerPrixNet

15      Exit Sub

AfficherErreur:

20      woups "frmCatalogueMec", "txtPrixNet_GotFocus", Err, Erl
End Sub

Private Sub CalculerPrixNet()

5       On Error GoTo AfficherErreur

10      Dim dblEscompte As Double
15      Dim dblPrix     As Double
20      Dim sPrix       As String
  
        'Si le prix net n'est pas barré.. ie.. si le prix spécial est vide
25      If txtPrixNet.Locked = False Then
30        mskEscompte.Text = Replace(mskEscompte.Text, "_", vbNullString)
    
35        mskEscompte.Text = Replace(mskEscompte.Text, ".", ",")
    
40        If mskEscompte.Text <> vbNullString Then
45          dblEscompte = CDbl(mskEscompte.Text)
50        Else
55          dblEscompte = 0
60        End If
          
65        If txtPrixList.Text <> vbNullString Then
70          sPrix = Replace(txtPrixList.Text, ".", ",")

75          If IsNumeric(sPrix) Then
80            dblPrix = CDbl(sPrix)
85          Else
90            dblPrix = 0
95          End If
100       Else
105         dblPrix = 0
110       End If
    
          'Calcul du prix net
115       txtPrixNet.Text = Round((dblPrix) * (1 - dblEscompte), 4)
    
120       txtPrixNet.Text = Replace(txtPrixNet.Text, ".", ",")
125     End If

130     Exit Sub

AfficherErreur:

135     woups "frmCatalogueMec", "CalculerPrixNet", Err, Erl
End Sub

Private Sub txtPrixNet_LostFocus()

5       On Error GoTo AfficherErreur

        'Vide le prix net si le user n'a rien marqué
10      If txtPrixNet.Text = "$0,00" Then
15        txtPrixNet.Text = vbNullString
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmCatalogueMec", "txtPrixNet_LostFocus", Err, Erl
End Sub

Private Sub mskValide_GotFocus()

5       On Error GoTo AfficherErreur

        'Si la date est sous le format AAAA-MM-JJ
10      If Len(mskValide.Text) = 10 Then
          'On la met sous le format AA-MM-JJ
15        mskValide.Text = Right$(mskValide.Text, 8)
20      End If
  
        'On met le mask
25      mskValide.mask = "##-##-##"

30      Exit Sub

AfficherErreur:

35      woups "frmCatalogueMec", "mskValide_GotFocus", Err, Erl
End Sub

Private Sub mskValide_LostFocus()

5       On Error GoTo AfficherErreur

        'On enlève le mask
10      mskValide.mask = vbNullString

15      If mskValide.Text = "__-__-__" Then
20        mskValide.Text = vbNullString
25      Else
30        If Len(mskValide.Text) = 8 Then
35          If IsDate(mskValide.Text) Then
              'On la met sous le format AAAA-MM-JJ
40            mskValide.Text = Year(DateSerial(Left$(mskValide.Text, 2), Mid$(mskValide.Text, 4, 2), Right$(mskValide.Text, 2))) & Mid$(mskValide.Text, 3, 8)
45          End If
50        End If
55      End If

60      Exit Sub

AfficherErreur:

65      woups "frmCatalogueMec", "mskValide_LostFocus", Err, Erl
End Sub

Private Sub cmbCategorie_Click()
        'pour sélectionner la bonne catégorie de pieces
  
5       On Error GoTo AfficherErreur

10      txtCategorie.Text = cmbCategorie.Text
  
15      Screen.MousePointer = vbHourglass

20      m_bBloqueDescription = True
 
25      m_bRempliManuel = True
  
30      Call cmbFabricant.Clear
  
35      Call cmbNoItem.Clear
  
40      Call ViderChamps_piece

45      Call RemplirComboFabricant
 
50      m_bBloqueDescription = False
  
55      Screen.MousePointer = vbDefault
  
60      Exit Sub

65      Exit Sub

AfficherErreur:

70      woups "FrmCatalogueMec", "cmbCategorie_Click", Err, Erl
End Sub

Public Sub RemplirComboCategorie()

15      On Error GoTo AfficherErreur

        'Remplir le combo categorie
5       Dim rstCatalogueMec As ADODB.Recordset
10      Dim iCompteur       As Integer
  
        'Il faut vider le combo avant de le remplir
20      Call cmbCategorie.Clear
    
        'Cette méthode crée un recordset contenant les categorie
        'le nom de toutes les tables de la BD
25      Set rstCatalogueMec = New ADODB.Recordset
       
30      Call rstCatalogueMec.Open("SELECT DISTINCT CATEGORIE FROM GRB_CatalogueMec ORDER BY CATEGORIE", g_connData, adOpenDynamic, adLockOptimistic)
      
        'Tant que ce n'est pas la fin des enregistrements
35      Do While Not rstCatalogueMec.EOF
40        If Not IsNull(rstCatalogueMec.Fields("CATEGORIE")) Then
45          Call cmbCategorie.AddItem(Trim(rstCatalogueMec.Fields("CATEGORIE")))
50        End If
    
55        Call rstCatalogueMec.MoveNext
60      Loop
  
65      Call rstCatalogueMec.Close
70      Set rstCatalogueMec = Nothing
  
        'Si le combo n'est pas vide, on sélectionne le premier
75      If cmbCategorie.ListCount > 0 Then
80        If m_sSelectCategorie <> "" Then
85          For iCompteur = 0 To cmbCategorie.ListCount - 1
90            If cmbCategorie.LIST(iCompteur) = m_sSelectCategorie Then
95              cmbCategorie.ListIndex = iCompteur

100             m_sSelectCategorie = ""

105             Exit For
110           End If
115         Next
120       Else
125         cmbCategorie.ListIndex = 0
130       End If
135     End If
  
140     Exit Sub
  
AfficherErreur:

145     woups "frmCatalogueMec", "RemplirComboCategorie", Err, Erl
End Sub

Private Sub RemplirComboDescription()

5       On Error GoTo AfficherErreur

        'Remplir le combo des descriptions
10      Dim rstCatMec  As ADODB.Recordset
15      Dim sPiece     As String
20      Dim sCategorie As String
21      Dim sFabricant As String

      
25      Do While m_collPieceDescFR.count > 0
30        Call m_collPieceDescFR.Remove(1)
35      Loop
  
40      Call cmbDescriptionFR.Clear
        
45      sCategorie = Replace(cmbCategorie.Text, "'", "''")
46      sFabricant = Replace(cmbFabricant.Text, "'", "''")
    
50      Set rstCatMec = New ADODB.Recordset
        
41      If sFabricant = "-- CHOISIR TOUS --" Then
            
            If cmbCategorie.Text = "Équipements et outillages - Sécurité/nettoyage" Or cmbCategorie.Text = "Quincaillerie - Boulons Hex. (Bolts) 1/4-20" Or cmbCategorie.Text = "Quincaillerie - Rondelle plate (Washer)" Or sChoisirTous = ")" Then

                    Call rstCatMec.Open("SELECT * FROM GRB_CatalogueMec WHERE CATEGORIE = '" & sCategorie & "' ORDER BY TRIM(PIECE)", g_connData, adOpenDynamic, adLockOptimistic)
54              Else
                    
                    Call rstCatMec.Open("SELECT * FROM GRB_CatalogueMec WHERE CATEGORIE = '" & sCategorie & "'" & sChoisirTous & " ORDER BY TRIM(PIECE)", g_connData, adOpenDynamic, adLockOptimistic)
55
                End If
        Else
                    
56          Call rstCatMec.Open("SELECT * FROM GRB_CatalogueMec WHERE CATEGORIE = '" & sCategorie & "' AND FABRICANT = '" & sFabricant & "' ORDER BY TRIM(PIECE)", g_connData, adOpenDynamic, adLockOptimistic)
57
        End If
        
  
60      Do While Not rstCatMec.EOF
65        If Not IsNull(rstCatMec.Fields("DESC_FR")) Then
70          If rstCatMec.Fields("DESC_FR") <> vbNullString Then
75            Call cmbDescriptionFR.AddItem(rstCatMec.Fields("DESC_FR"))
      
80            sPiece = rstCatMec.Fields("PIECE")
     
85            Call m_collPieceDescFR.Add(sPiece)
90          End If
95        End If
    
100       Call rstCatMec.MoveNext
105     Loop
  
110     Call rstCatMec.Close
115     Set rstCatMec = Nothing

120     Exit Sub

AfficherErreur:

125     woups "frmCatalogueMec", "RemplirComboDescription", Err, Erl
End Sub

Private Sub RemplirComboFRS()

5       On Error GoTo AfficherErreur

        'Remplir le combo des fournisseurs
10      Dim rstPieceFRS As ADODB.Recordset
  
        'Il faut vider le combo avant de le remplir
15      Call cmbfrs.Clear
      
        'ouvre la table grb_Fournisseur
20      Set rstPieceFRS = New ADODB.Recordset
        
25      Call rstPieceFRS.Open("SELECT * FROM GRB_Fournisseur WHERE Supprimé = False ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)

        'Tant que ce n'est pas la fin des enregistrements
30      Do While Not rstPieceFRS.EOF
35        Call cmbfrs.AddItem(rstPieceFRS.Fields("NomFournisseur"))
40        cmbfrs.ItemData(cmbfrs.newIndex) = rstPieceFRS.Fields("IDFRS")
    
45        Call rstPieceFRS.MoveNext
50      Loop
  
55      Call rstPieceFRS.Close
60      Set rstPieceFRS = Nothing

65      Exit Sub

AfficherErreur:

70      woups "frmCatalogueMec", "RemplirComboFRS", Err, Erl
End Sub

Private Sub txtPrixSpecial_LostFocus()

5       On Error GoTo AfficherErreur
  
10      txtPrixSpecial.Text = Replace(txtPrixSpecial.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmCatalogueMec", "txtPrixSpecial_LostFocus", Err, Erl
End Sub

Private Sub RemplirComboLocalisation()

5       On Error GoTo AfficherErreur

        'Rempli le combo cmbLocalisation
10      Dim rstLocalisation As ADODB.Recordset
  
15      Set rstLocalisation = New ADODB.Recordset
  
20      Call rstLocalisation.Open("SELECT DISTINCT Localisation FROM GRB_InventaireMec", g_connData, adOpenDynamic, adLockOptimistic)
  
        'Il faut vider le combo avant de le remplir
25      Call cmbLocalisation.Clear
  
        'Tant que ce n'est pas la fin des enregistrements
30      Do While Not rstLocalisation.EOF
          'Si l'enregistrement n'est pas Null
35        If Not IsNull(rstLocalisation.Fields("Localisation")) Then
40          If Trim(rstLocalisation.Fields("Localisation")) <> "" Then
              'On l'ajoute dans le combo
45            Call cmbLocalisation.AddItem(rstLocalisation.Fields("Localisation"))
50          End If
55        End If
    
60        Call rstLocalisation.MoveNext
65      Loop
  
70      Call rstLocalisation.Close
75      Set rstLocalisation = Nothing

80      Exit Sub

AfficherErreur:

85      woups "frmCatalogueMec", "RemplirComboLocalisation", Err, Erl
End Sub

Private Sub txtQuantitéBoite_LostFocus()

5       On Error GoTo AfficherErreur

10      txtQuantitéBoite.Text = Replace(txtQuantitéBoite.Text, ".", ",")

15      If Not IsNumeric(txtQuantitéBoite.Text) Or txtQuantitéBoite.Text = "0" Then
20        txtQuantitéBoite.Text = "1"
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmCatalogueMec", "txtQuantitéBoite_LostFocus", Err, Erl
End Sub

Private Sub txtQuantiteCommande_LostFocus()

5       On Error GoTo AfficherErreur

10      txtQuantiteCommande.Text = Replace(txtQuantiteCommande.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmCatalogueMec", "txtQuantiteCommande_LostFocus", Err, Erl
End Sub

Private Sub txtQuantiteMinimum_LostFocus()

5       On Error GoTo AfficherErreur

10      txtQuantiteMinimum.Text = Replace(txtQuantiteMinimum.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmCatalogueMec", "txtQuantiteMinimum_LostFocus", Err, Erl
End Sub

Private Sub txtQuantiteStock_LostFocus()

5       On Error GoTo AfficherErreur

10      txtQuantiteStock.Text = Replace(txtQuantiteStock.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmCatalogueMec", "txtQuantiteStock_LostFocus", Err, Erl
End Sub
