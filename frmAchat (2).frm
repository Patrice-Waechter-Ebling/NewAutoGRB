VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAchat 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13410
   Icon            =   "frmAchat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7425
   ScaleWidth      =   13410
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
      Left            =   4440
      TabIndex        =   60
      Top             =   4200
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton cmdOKDateRequise 
         Caption         =   "OK"
         Height          =   375
         Left            =   3480
         TabIndex        =   62
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdAnnulerDateRequise 
         Caption         =   "Annuler"
         Height          =   375
         Left            =   3480
         TabIndex        =   61
         Top             =   1200
         Width           =   1095
      End
      Begin MSComCtl2.MonthView mvwDateRequise 
         Height          =   2370
         Left            =   600
         TabIndex        =   63
         Top             =   360
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   0
         Appearance      =   1
         StartOfWeek     =   152633345
         CurrentDate     =   38247
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
      Left            =   1560
      TabIndex        =   13
      Top             =   1200
      Visible         =   0   'False
      Width           =   10335
      Begin VB.CommandButton cmdOKPieceTrouve 
         Caption         =   "OK"
         Height          =   375
         Left            =   9120
         TabIndex        =   16
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdAnnulerPieceTrouve 
         Caption         =   "Annuler"
         Height          =   375
         Left            =   7920
         TabIndex        =   15
         Top             =   2280
         Width           =   1095
      End
      Begin MSComctlLib.ListView lvwPieceTrouve 
         Height          =   1935
         Left            =   120
         TabIndex        =   14
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
      TabIndex        =   44
      Top             =   4440
      Visible         =   0   'False
      Width           =   8895
      Begin VB.CommandButton cmdAnnulerPrix 
         Caption         =   "Annuler"
         Height          =   375
         Left            =   6240
         TabIndex        =   53
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdOKPrix 
         Caption         =   "OK"
         Height          =   375
         Left            =   7440
         TabIndex        =   52
         Top             =   1800
         Width           =   1095
      End
      Begin VB.OptionButton optSpain 
         BackColor       =   &H00000000&
         Caption         =   "SPA"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   8040
         TabIndex        =   51
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optCAN 
         BackColor       =   &H00000000&
         Caption         =   "CAN"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6600
         TabIndex        =   50
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
         TabIndex        =   49
         Top             =   240
         Width           =   735
      End
      Begin VB.ComboBox cmbfrs 
         Height          =   315
         Left            =   240
         TabIndex        =   48
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox txtPrixNet 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4920
         TabIndex        =   47
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtPrixList 
         BackColor       =   &H00FFFFFF&
         DataField       =   "PRIX_LIST"
         DataSource      =   "DatCat1"
         Height          =   285
         Left            =   4920
         TabIndex        =   46
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtPrixSpecial 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4920
         TabIndex        =   45
         Top             =   1560
         Width           =   855
      End
      Begin MSMask.MaskEdBox mskEscompte 
         Height          =   255
         Left            =   4920
         TabIndex        =   54
         Top             =   840
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Image imgEU 
         Height          =   1065
         Left            =   6840
         Picture         =   "frmAchat.frx":030A
         Stretch         =   -1  'True
         Top             =   600
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Image imgSpain 
         Height          =   1065
         Left            =   6840
         Picture         =   "frmAchat.frx":4D07C
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
         Index           =   5
         Left            =   3720
         TabIndex        =   59
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
         Index           =   4
         Left            =   3720
         TabIndex        =   58
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
         Index           =   3
         Left            =   3720
         TabIndex        =   57
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
         Index           =   2
         Left            =   240
         TabIndex        =   56
         Top             =   240
         Width           =   1215
      End
      Begin VB.Image imgCanada 
         Height          =   1065
         Left            =   6840
         Picture         =   "frmAchat.frx":4F50B
         Stretch         =   -1  'True
         Top             =   600
         Visible         =   0   'False
         Width           =   1680
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
         Index           =   1
         Left            =   3720
         TabIndex        =   55
         Top             =   1560
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdReception 
      Caption         =   "Réception"
      Height          =   375
      Left            =   2640
      TabIndex        =   43
      Top             =   6960
      Width           =   975
   End
   Begin VB.Frame fraInventaire 
      BackColor       =   &H00000000&
      Caption         =   "Inventaire à commander"
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
      Height          =   3135
      Left            =   120
      TabIndex        =   18
      Top             =   1200
      Visible         =   0   'False
      Width           =   13215
      Begin VB.CommandButton cmdAnnulerInventaire 
         Caption         =   "Annuler"
         Height          =   375
         Left            =   10560
         TabIndex        =   20
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CommandButton cmdOKInventaire 
         Caption         =   "OK"
         Height          =   375
         Left            =   11880
         TabIndex        =   21
         Top             =   2640
         Width           =   1215
      End
      Begin MSComctlLib.ListView lvwInventaire 
         Height          =   2295
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   4048
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No. d'item"
            Object.Width           =   3254
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Manufacturier"
            Object.Width           =   2037
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description"
            Object.Width           =   7144
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Commentaire"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Qté en stock"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Qté minimum"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Qté à commander"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.CommandButton cmdInventaire 
      Caption         =   "Inventaire à commander"
      Height          =   375
      Left            =   11280
      TabIndex        =   24
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   1440
      TabIndex        =   31
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton cmdRetour 
      Caption         =   "Retour"
      Height          =   375
      Left            =   3720
      TabIndex        =   32
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton cmdDemande 
      Caption         =   "Demande de prix"
      Height          =   375
      Left            =   4800
      TabIndex        =   33
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton cmdTri 
      Caption         =   "Trier"
      Height          =   315
      Left            =   7440
      TabIndex        =   25
      Top             =   1380
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
      Left            =   9000
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "0"
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdRafraichir 
      Caption         =   "Rafraichir"
      Height          =   315
      Left            =   7440
      TabIndex        =   12
      Top             =   1080
      Width           =   975
   End
   Begin VB.ComboBox cmbTri 
      Height          =   315
      ItemData        =   "frmAchat.frx":A54ED
      Left            =   5640
      List            =   "frmAchat.frx":A5500
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmdFermer 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   12120
      TabIndex        =   42
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdAjouter 
      Caption         =   "Ajouter"
      Height          =   375
      Left            =   8160
      TabIndex        =   36
      Top             =   6960
      Width           =   1215
   End
   Begin VB.ComboBox cmbCategorie 
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
      ItemData        =   "frmAchat.frx":A5557
      Left            =   120
      List            =   "frmAchat.frx":A5559
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   1320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtRaison 
      Height          =   285
      Left            =   4560
      MaxLength       =   255
      TabIndex        =   8
      Top             =   480
      Width           =   5775
   End
   Begin VB.ComboBox cmbNoAchat 
      Height          =   315
      Left            =   4560
      TabIndex        =   1
      Top             =   120
      Width           =   2295
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
      Height          =   1815
      Left            =   1320
      TabIndex        =   27
      Top             =   1800
      Visible         =   0   'False
      Width           =   10935
      Begin MSComctlLib.ListView lvwFournisseur 
         Height          =   1455
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   2566
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
      End
   End
   Begin VB.CommandButton cmdModifier 
      Caption         =   "Modifier"
      Height          =   375
      Left            =   10800
      TabIndex        =   41
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdSupprimer 
      Caption         =   "Supprimer"
      Height          =   375
      Left            =   9480
      TabIndex        =   39
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprimer 
      Caption         =   "Imprimer"
      Height          =   375
      Left            =   120
      TabIndex        =   30
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox txtAcheteur 
      Height          =   288
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdBonCommande 
      Caption         =   "Bon de commande"
      Height          =   375
      Left            =   6480
      TabIndex        =   34
      Top             =   6960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSComctlLib.ListView lvwAchat 
      Height          =   2535
      Left            =   120
      TabIndex        =   29
      Top             =   4320
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   4471
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
         Text            =   "Qté"
         Object.Width           =   794
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "No. d'item"
         Object.Width           =   3228
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   6720
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Manufacturier"
         Object.Width           =   2037
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Prix listé"
         Object.Width           =   1623
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Escompte"
         Object.Width           =   1561
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Prix net"
         Object.Width           =   1623
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Distributeur"
         Object.Width           =   1773
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "TOTAL"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Date Commande"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Date Requise"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvwPieces 
      Height          =   2535
      Left            =   120
      TabIndex        =   26
      Top             =   1680
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   4471
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
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Commentaire"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtNoAchat 
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtDate 
      Height          =   288
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdEnregistrer 
      Caption         =   "Enregistrer"
      Height          =   375
      Left            =   9480
      TabIndex        =   38
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdAnnuler 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   10800
      TabIndex        =   40
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdMaterielInutile 
      Caption         =   "Inutile"
      Height          =   375
      Left            =   8160
      TabIndex        =   37
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdMauvaisPrix 
      Caption         =   "Prix"
      Height          =   375
      Left            =   6840
      TabIndex        =   35
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Date                    AA-MM-JJ"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3840
      TabIndex        =   9
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblPrixTotal 
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
      Left            =   8160
      TabIndex        =   5
      Top             =   840
      Width           =   750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Numéro"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   3840
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblTri 
      BackStyle       =   0  'Transparent
      Caption         =   "Trier par :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5640
      TabIndex        =   17
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblRaison 
      BackStyle       =   0  'Transparent
      Caption         =   "Raison "
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblCategorie 
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
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblNoSoumission 
      BackStyle       =   0  'Transparent
      Caption         =   "Acheteur"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7680
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.Menu mnuRightClick 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuDateRequise 
         Caption         =   "Modifier la date requise"
      End
   End
End
Attribute VB_Name = "frmAchat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Index des colonnes de lvwAchat
'Énumération servant à savoir si le form est en mode modif/ajout ou en mode
'inactif (affichage seulement)
Private Enum enumMode
 MODE_AJOUT_MODIF = 0
 MODE_INACTIF = 1
End Enum

'Pour la recherche de pièce dans lvwPieces

Private m_bMonthViewHasFocus As Boolean

Private Sub AnnulerCommande()

 On Error GoTo Oups

 Dim itmAvant As ListItem
 Dim itmAnnulation As ListItem

 Set itmAvant = lvwAchat.SelectedItem
 Set itmAnnulation = lvwAchat.ListItems.Add(itmAvant.Index + 1)

 itmAnnulation.Checked = itmAvant.Checked

 'Quantité
 itmAnnulation.Text = "-" & itmAvant.Text

 'On met l'id de la section dans le tag du listItem
 itmAnnulation.Tag = itmAvant.Tag

 'No d'item
 itmAnnulation.SubItems(I_COL_ACHAT_PIECE) = itmAvant.SubItems(I_COL_ACHAT_PIECE)

 'On met le nom de la sous-section dans le tag du no d'item
 itmAnnulation.ListSubItems(I_COL_ACHAT_PIECE).Tag = itmAvant.ListSubItems(I_COL_ACHAT_PIECE).Tag

 'On met la description en francais dans la colonne et la description en anglais
 'dans le tag
 itmAnnulation.SubItems(I_COL_ACHAT_DESCR) = itmAvant.SubItems(I_COL_ACHAT_DESCR)
  itmAnnulation.ListSubItems(I_COL_ACHAT_DESCR).Tag = itmAvant.ListSubItems(I_COL_ACHAT_DESCR).Tag

 'On met le nom du fabricant dans la col-nne et l'ordre de la section dans le tag
  itmAnnulation.SubItems(I_COL_ACHAT_MANUFACT) = itmAvant.SubItems(I_COL_ACHAT_MANUFACT)
  itmAnnulation.ListSubItems(I_COL_ACHAT_MANUFACT).Tag = itmAvant.ListSubItems(I_COL_ACHAT_MANUFACT).Tag

 'Prix listé
  itmAnnulation.SubItems(I_COL_ACHAT_PRIX_LIST) = itmAvant.SubItems(I_COL_ACHAT_PRIX_LIST)

  itmAnnulation.ListSubItems(I_COL_ACHAT_PRIX_LIST).Tag = itmAvant.ListSubItems(I_COL_ACHAT_PRIX_LIST).Tag

  itmAnnulation.SubItems(I_COL_ACHAT_ESCOMPTE) = itmAvant.SubItems(I_COL_ACHAT_ESCOMPTE)

  itmAnnulation.SubItems(I_COL_ACHAT_PRIX_NET) = itmAvant.SubItems(I_COL_ACHAT_PRIX_NET)

 'On met le fournisseur dans la colonne et l'id dans le tag
  itmAnnulation.SubItems(I_COL_ACHAT_DISTRIB) = itmAvant.SubItems(I_COL_ACHAT_DISTRIB)
10 itmAnnulation.ListSubItems(I_COL_ACHAT_DISTRIB).Tag = itmAvant.ListSubItems(I_COL_ACHAT_DISTRIB).Tag

 'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
itmAnnulation.SubItems(I_COL_ACHAT_TOTAL) = "-" & itmAvant.SubItems(I_COL_ACHAT_TOTAL)

itmAnnulation.ForeColor = COLOR_VERT_FORET
itmAnnulation.ListSubItems(I_COL_ACHAT_PIECE).ForeColor = COLOR_VERT_FORET
itmAnnulation.ListSubItems(I_COL_ACHAT_DESCR).ForeColor = COLOR_VERT_FORET
itmAnnulation.ListSubItems(I_COL_ACHAT_DISTRIB).ForeColor = COLOR_VERT_FORET
itmAnnulation.ListSubItems(I_COL_ACHAT_ESCOMPTE).ForeColor = COLOR_VERT_FORET
itmAnnulation.ListSubItems(I_COL_ACHAT_MANUFACT).ForeColor = COLOR_VERT_FORET
itmAnnulation.ListSubItems(I_COL_ACHAT_PRIX_LIST).ForeColor = COLOR_VERT_FORET
itmAnnulation.ListSubItems(I_COL_ACHAT_PRIX_NET).ForeColor = COLOR_VERT_FORET
itmAnnulation.ListSubItems(I_COL_ACHAT_TOTAL).ForeColor = COLOR_VERT_FORET
 
itmAnnulation.Bold = True
1  itmAnnulation.ListSubItems(I_COL_ACHAT_PIECE).Bold = True
itmAnnulation.ListSubItems(I_COL_ACHAT_DESCR).Bold = True
 itmAnnulation.ListSubItems(I_COL_ACHAT_DISTRIB).Bold = True
itmAnnulation.ListSubItems(I_COL_ACHAT_ESCOMPTE).Bold = True
 itmAnnulation.ListSubItems(I_COL_ACHAT_MANUFACT).Bold = True
itmAnnulation.ListSubItems(I_COL_ACHAT_PRIX_LIST).Bold = True
 itmAnnulation.ListSubItems(I_COL_ACHAT_PRIX_NET).Bold = True
1  itmAnnulation.ListSubItems(I_COL_ACHAT_TOTAL).Bold = True

 itmAvant.ForeColor = COLOR_NOIR
 itmAvant.ListSubItems(I_COL_ACHAT_PIECE).ForeColor = COLOR_NOIR
itmAvant.ListSubItems(I_COL_ACHAT_DESCR).ForeColor = COLOR_NOIR
itmAvant.ListSubItems(I_COL_ACHAT_DISTRIB).ForeColor = COLOR_NOIR
itmAvant.ListSubItems(I_COL_ACHAT_ESCOMPTE).ForeColor = COLOR_NOIR
itmAvant.ListSubItems(I_COL_ACHAT_MANUFACT).ForeColor = COLOR_NOIR
itmAvant.ListSubItems(I_COL_ACHAT_PRIX_LIST).ForeColor = COLOR_NOIR
itmAvant.ListSubItems(I_COL_ACHAT_PRIX_NET).ForeColor = COLOR_NOIR
itmAvant.ListSubItems(I_COL_ACHAT_TOTAL).ForeColor = COLOR_NOIR
itmAvant.ListSubItems(I_COL_ACHAT_DATE_COMMANDE).ForeColor = COLOR_NOIR
itmAvant.ListSubItems(I_COL_ACHAT_DATE_REQUISE).ForeColor = COLOR_NOIR

Call lvwAchat.Refresh

2  Call CalculerPrix

Exit Sub

Oups:

2  wOups "frmAchat", "AnnulerCommande", Err, Err.number, Err.Description
End Sub

Public Sub Afficher(ByVal eCatalogue As enumCatalogue)

 On Error GoTo Oups

 Call Unload(frmChoixProjSoum)
 
 m_eCatalogue = eCatalogue
 
 Select Case eCatalogue
 'Si c'est électrique
 Case ELECTRIQUE:
 Me.Caption = "Achat électrique"
 cmbCategorie.width = I_WIDTH_CATEGORIE_ELEC
 
 Case MECANIQUE:
 Me.Caption = "Achat mécanique"
 cmbCategorie.width = I_WIDTH_CATEGORIE_MEC
 End Select
 
 'Initialise le tri à PIECE_GRB
 cmbTri.ListIndex = I_CMB_PIECE
 
 Call RemplirComboAchat(vbNullString)
 
 'Rempli le combo des catégories de pièce
  Call RemplirComboCategorie
 
  Call AfficherControles(MODE_INACTIF)
 
  Call Me.Show

  Exit Sub

Oups:

  wOups "frmAchat", "Afficher", Err, Err.number, Err.Description
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
 Dim bBonCommande As Boolean
 Dim bTri As Boolean
 Dim bCmbAchat As Boolean
  Dim bDemandePrix As Boolean
  Dim bRetour As Boolean
  Dim bInventaire As Boolean
  Dim bInutile As Boolean
  Dim bPrix As Boolean
  Dim bPieces As Boolean
  Dim bReception As Boolean
 
  m_eMode = eMode
 
10 Select Case eMode
 Case MODE_AJOUT_MODIF:
bEnregistrer = True
 bAnnuler = True
 bPieces = True
 bTri = True
 bInventaire = True
 bPrix = True
 bInutile = True
 
 Case MODE_INACTIF:
 bModifier = True
 bFermer = True
 bImprimer = True
 bAjouter = True
 bBonCommande = True
 bSupprimer = True
 bCmbAchat = True
 bDemandePrix = True
 bRetour = True
 bReception = True
 End Select
 
1  cmbNoAchat.Visible = bCmbAchat
 txtNoAchat.Visible = Not bCmbAchat
 
 Cmdajouter.Visible = bAjouter
cmdModifier.Visible = bModifier
cmdsupprimer.Visible = bSupprimer
cmdEnregistrer.Visible = bEnregistrer
cmdAnnuler.Visible = bAnnuler
Cmdfermer.Visible = bFermer
cmdImprimer.Visible = bImprimer
cmdBonCommande.Visible = bBonCommande
cmdDemande.Visible = bDemandePrix
cmdRetour.Visible = bRetour
cmdInventaire.Visible = bInventaire
2  cmdReception.Visible = bReception
lblCategorie.Visible = bPieces
2  cmbCategorie.Visible = bPieces
lvwPieces.Visible = bPieces

2  lblTri.Visible = bTri
cmbTri.Visible = bTri
2  cmdTri.Visible = bTri
cmdRafraichir.Visible = bTri
 
 'Exception puisqu'il y en a qu'un seul
30 If m_eMode = MODE_AJOUT_MODIF Then
3 txtRaison.Locked = False
Else
 txtRaison.Locked = True
End If

If m_eMode = MODE_AJOUT_MODIF Then
 lvwAchat.Top = I_TOP_AJOUT_MODIF
 lvwAchat.Height = I_HEIGHT_AJOUT_MODIF
Else
 lvwAchat.Top = I_TOP_INACTIF
 lvwAchat.Height = I_HEIGHT_INACTIF
End If

3  Exit Sub

Oups:

wOups "frmAchat", "AfficherControles", Err, Err.number, Err.Description
End Sub

Private Sub cmbCategorie_Click()

 On Error GoTo Oups

 'Rempli lvwPieces selon la catégorie de pièce choisie
 Call RemplirListViewPieces

 Exit Sub

Oups:

 wOups "frmAchat", "cmbCategorie_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmbNoAchat_Click()

 On Error GoTo Oups

 Dim sNomClient As String
 Dim sNomContact As String
 Dim iCompteur As Integer
 Dim rstAchat As ADODB.Recordset
 Dim sIDAchat As String
 Dim iIndexAchat As Integer
 
 Screen.MousePointer = vbHourglass
 
 txtNoAchat.Text = cmbNoAchat.Text
 
 'Rempli les valeurs de l'achat sélectionné
 Call RemplirAchat

 sIDAchat = Trim$(Left$(txtNoAchat.Text, InStr(1, txtNoAchat.Text, "-") + 2))
  iIndexAchat = CInt(Trim$(Right$(txtNoAchat.Text, Len(txtNoAchat.Text) - InStrRev(txtNoAchat.Text, "-"))))

  Set rstAchat = New ADODB.Recordset

  Call rstAchat.Open("SELECT Modification, Par FROM GrbAchat WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat, g_connData, adOpenDynamic, adLockOptimistic)

  If rstAchat.Fields("Modification") = True And rstAchat.Fields("Par") = g_sEmploye Then
  cmdReset.Visible = True
  Else
  cmdReset.Visible = False
  End If

10 Call rstAchat.Close
Set rstAchat = Nothing
 
Screen.MousePointer = vbDefault

Exit Sub

Oups:

wOups "frmAchat", "cmbNoAchat_Click", Err, Err.number, Err.Description
End Sub

Private Sub RemplirAchat()

 On Error GoTo Oups

 Dim rstAchat As ADODB.Recordset
 Dim rstEmploye As ADODB.Recordset
 Dim sNoAchat As String
 Dim iIndexAchat As Integer
 
 sNoAchat = Left$(txtNoAchat.Text, 9)
 
 iIndexAchat = CInt(Right$(txtNoAchat.Text, 3))
 
 Set rstAchat = New ADODB.Recordset
 Set rstEmploye = New ADODB.Recordset
 
 Call rstAchat.Open("SELECT * FROM GrbAchat WHERE IDAchat = '" & sNoAchat & "' AND IndexAchat = " & iIndexAchat, g_connData, adOpenDynamic, adLockOptimistic)
 
 Call rstEmploye.Open("SELECT Employe FROM GrbEmployés WHERE noEmploye = " & rstAchat.Fields("Acheteur"), g_connData, adOpenDynamic, adLockOptimistic)
 
  txtAcheteur.Text = rstEmploye.Fields("employe")
  txtAcheteur.Tag = rstAchat.Fields("Acheteur")
 
  Call rstEmploye.Close
  Set rstEmploye = Nothing
 
  txtRaison.Text = rstAchat.Fields("Raison")
  txtDate.Text = rstAchat.Fields("DateAchat")
 
  txtPrixTotal.Text = Conversion(rstAchat.Fields("PrixTotal"), MODE_ARGENT)

  Call rstAchat.Close
10 Set rstAchat = Nothing
 
Call RemplirListViewAchat

Exit Sub

Oups:

wOups "frmAchat", "RemplirAchat", Err, Err.number, Err.Description
End Sub

Private Sub cmbNoAchat_KeyUp(KeyCode As Integer, Shift As Integer)
 
 On Error GoTo Oups

 Dim iCompteur As Integer
 
 For iCompteur = 0 To cmbNoAchat.ListCount - 1
 If UCase(cmbNoAchat.LIST(iCompteur)) = UCase(cmbNoAchat.Text) Then
 cmbNoAchat.ListIndex = iCompteur
 
 Exit For
 End If
 Next

 Exit Sub

Oups:

 wOups "frmAchat", "cmbNoAchat_KeyUp", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnuler_Click()

 On Error GoTo Oups

 Dim sChamps As String
 Dim sTable As String
 Dim sTablePiece As String
 
 Screen.MousePointer = vbHourglass
 
 'Initialisation des variables booléennes
 m_bInventaire = False
 m_bMauvaisPrix = False
 m_bPieceInutile = False
 m_bRecherchePiece = False
 
 'Remet en mode inactif
 Call AfficherControles(MODE_INACTIF)
 
 Call OuvrirAchat(False)
 
  Call RemplirComboAchat(m_sAncienAchat)
 
  m_bModeAjout = False
 
  Screen.MousePointer = vbDefault

  Exit Sub

Oups:

  wOups "frmAchat", "cmdAnnuler_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnulerDateRequise_Click()

 On Error GoTo Oups

 fraDateRequise.Visible = False

 m_bMonthViewHasFocus = False

 Exit Sub

Oups:

 wOups "frmAchat", "cmdAnnulerDateRequise_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnulerDateRequise_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
 On Error GoTo Oups

 If m_bMonthViewHasFocus = True Then
 Call cmdAnnulerDateRequise_Click
 End If

 Exit Sub

Oups:

 wOups "frmAchat", "cmdAnnulerDateRequise_MouseUp", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnulerInventaire_Click()

 On Error GoTo Oups

 fraInventaire.Visible = False

 Exit Sub

Oups:

 wOups "frmAchat", "cmdAnnulerInventaire_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdBonCommande_Click()

 On Error GoTo Oups

 Dim rstAchat As ADODB.Recordset

 Set rstAchat = New ADODB.Recordset

 Call rstAchat.Open("SELECT Modification, Par FROM GrbAchat WHERE IDAchat = '" & Left$(txtNoAchat.Text, 9) & "' AND IndexAchat = " & CInt(Right$(txtNoAchat.Text, 3)), g_connData, adOpenDynamic, adLockOptimistic)

 If rstAchat.Fields("Modification") = False Then
 If lvwAchat.ListItems.count > 0 Then
 If m_eCatalogue = ELECTRIQUE Then
 Call frmChoixBonCommande.AfficherAchat(Left$(txtNoAchat.Text, 9), CInt(Right$(txtNoAchat.Text, 3)), ELECTRIQUE)
 Else
 Call frmChoixBonCommande.AfficherAchat(Left$(txtNoAchat.Text, 9), CInt(Right$(txtNoAchat.Text, 3)), MECANIQUE)
 End If
  Else
  Call MsgBox("Il n'y a pas de pièces à commander pour cet achat!", vbOKOnly, "Erreur")
  End If
  Else
  Call MsgBox("Ce projet est en modification par " & rstAchat.Fields("Par") & "!", vbOKOnly, "Erreur")
  End If

  Call rstAchat.Close
  Set rstAchat = Nothing

10 Exit Sub

Oups:

wOups "frmAchat", "cmdBonCommande_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdDemande_Click()

 On Error GoTo Oups

 Dim rstAchat As ADODB.Recordset
 Dim sIDAchat As String
 Dim iIndexAchat As Integer

 sIDAchat = Trim$(Left$(txtNoAchat.Text, InStr(1, txtNoAchat.Text, "-") + 2))
 iIndexAchat = CInt(Trim$(Right$(txtNoAchat.Text, Len(txtNoAchat.Text) - InStrRev(txtNoAchat.Text, "-"))))

 Set rstAchat = New ADODB.Recordset

 Call rstAchat.Open("SELECT Modification, Par FROM GrbAchat WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat, g_connData, adOpenDynamic, adLockOptimistic)

 If rstAchat.Fields("Modification") = False Then
 Call frmChoixDemande.AfficherAchat(txtNoAchat.Text, m_eCatalogue, MODE_PIECE)
 Else
  Call MsgBox("Cet achat est en modification par " & rstAchat.Fields("Par") & "!", vbOKOnly, "Erreur")
  End If

  Call rstAchat.Close
  Set rstAchat = Nothing
 
  Exit Sub

Oups:

  wOups "frmAchat", "cmdDemande_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdImprimer_Click()

 On Error GoTo Oups

 Dim rstAchat As ADODB.Recordset
 Dim sIDAchat As String
 Dim iIndexAchat As Integer

 sIDAchat = Trim$(Left$(txtNoAchat.Text, InStr(1, txtNoAchat.Text, "-") + 2))
 iIndexAchat = CInt(Trim$(Right$(txtNoAchat.Text, Len(txtNoAchat.Text) - InStrRev(txtNoAchat.Text, "-"))))

 Set rstAchat = New ADODB.Recordset

 Call rstAchat.Open("SELECT Modification, Par FROM GrbAchat WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat, g_connData, adOpenDynamic, adLockOptimistic)

 If rstAchat.Fields("Modification") = False Then
 Call frmChoixImpressionAchat.Afficher(m_eCatalogue)
 Else
  Call MsgBox("Cet achat est en modification par " & rstAchat.Fields("Par") & "!", vbOKOnly, "Erreur")
  End If

  Call rstAchat.Close
  Set rstAchat = Nothing

  Exit Sub

Oups:

  wOups "frmAchat", "cmdImprimer_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdInventaire_Click()

 On Error GoTo Oups

 Call RemplirListViewInventaire

 fraInventaire.Visible = True

 Exit Sub

Oups:

 wOups "frmAchat", "cmdInventaire_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdOKDateRequise_Click()

 On Error GoTo Oups

 lvwAchat.SelectedItem.SubItems(I_COL_ACHAT_DATE_REQUISE) = ConvertDate(mvwDateRequise.Value)

 lvwAchat.SelectedItem.ListSubItems(I_COL_ACHAT_DATE_REQUISE).ForeColor = COLOR_ORANGE

 fraDateRequise.Visible = False

 m_bMonthViewHasFocus = False

 Exit Sub

Oups:

 wOups "frmAchat", "cmdOKDateRequise_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdOKDateRequise_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
 On Error GoTo Oups

 If m_bMonthViewHasFocus = True Then
 Call cmdOKDateRequise_Click
 End If

 Exit Sub

Oups:

 wOups "frmAchat", "cmdOKDateRequise_MouseUp", Err, Err.number, Err.Description
End Sub

Private Sub cmdOKInventaire_Click()

 On Error GoTo Oups

 'On affiche lvwFournisseur selon la pièce choisie
 'Rempli les fournisseurs de la pièce choisie
 m_bInventaire = True
 m_bRecherchePiece = False
 
 Call AfficherListeFournisseurs
 
 'Si le listview n'est pas vide
 If lvwfournisseur.ListItems.count = 1 Then
 If MsgBox("Il n'y a aucun fournisseur pour cette pièce!" & vbNewLine & "Voulez-vous en ajouter?", vbYesNo, "Erreur") = vbYes Then
 Screen.MousePointer = vbHourglass
 
 'On ouvre le catalogue sur cet enregistrement
 If m_eCatalogue = ELECTRIQUE Then
 Call FrmCatalogueElec.AfficherForm(cmbCategorie.Text, lvwInventaire.SelectedItem.SubItems(I_COL_INV_MANUFACT), lvwInventaire.SelectedItem.Text)
 Else
 Call FrmCatalogueMec.AfficherForm(cmbCategorie.Text, lvwInventaire.SelectedItem.SubItems(I_COL_INV_MANUFACT), lvwInventaire.SelectedItem.Text)
  End If
 
  Screen.MousePointer = vbDefault
 
 'On rappelle la méthode
  Call AfficherListeFournisseurs
  End If
  End If

  fraInventaire.Visible = False

  Exit Sub

Oups:

  wOups "frmAchat", "cmdOKInventaire_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdRafraichir_Click()

 On Error GoTo Oups

 If m_sTri <> vbNullString Then
 m_sTri = vbNullString
 
 Call RemplirListViewPieces
 End If

 Exit Sub

Oups:

 wOups "frmAchat", "cmdRafraichir_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdReset_Click()
 'Permet d'effacer le champs Modification et Par si c'est le user actuel
 On Error GoTo Oups

 Dim rstAchat As ADODB.Recordset
 Dim sIDAchat As String
 Dim iIndexAchat As Integer

 If MsgBox("Êtes-vous certains de ne pas être en modification sur un autre ordinateur?", vbYesNo) = vbYes Then
 sIDAchat = Trim$(Left$(txtNoAchat.Text, InStr(1, txtNoAchat.Text, "-") + 2))
 iIndexAchat = CInt(Trim$(Right$(txtNoAchat.Text, Len(txtNoAchat.Text) - InStrRev(txtNoAchat.Text, "-"))))

 Set rstAchat = New ADODB.Recordset

 Call rstAchat.Open("SELECT Modification, Par FROM GrbAchat WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat, g_connData, adOpenDynamic, adLockOptimistic)

 rstAchat.Fields("Modification") = False
 rstAchat.Fields("Par") = ""

  Call rstAchat.Update

  Call rstAchat.Close
  Set rstAchat = Nothing

  cmdReset.Visible = False
  End If

  Exit Sub

Oups:

  wOups "frmAchat", "cmdReset_Click", Err, Err.number, Err.Description
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

 wOups "frmAchat", "cmdTri_Click", Err, Err.number, Err.Description
End Sub

Private Sub lvwAchat_DblClick()

 On Error GoTo Oups

 'Si en mode ajout ou modif
 If m_eMode = MODE_AJOUT_MODIF Then
 'Si la liste n'est pas vide
 If lvwAchat.ListItems.count > 0 Then
 'Si la pièce n'a pas de fournisseur
 If Trim$(lvwAchat.SelectedItem.SubItems(I_COL_ACHAT_DISTRIB)) = vbNullString Then
 Call ViderChamps_frs

 cmbfrs.Locked = False

 m_bMauvaisPrix = False

 'Rempli le combo des fournisseurs
 Call RemplirComboFournisseur
 
 'Montre le frame
 fraPrixPiece.Visible = True

 'Met le numéro de la pièce dans le tag
 fraPrixPiece.Tag = lvwAchat.SelectedItem.Index
 
 'Donne le focus au combo
 Call cmbfrs.SetFocus
  Else
  'Si le listItem est orange
  If lvwAchat.SelectedItem.ListSubItems(I_COL_ACHAT_PIECE).ForeColor = COLOR_ORANGE Then
  If MsgBox("Voulez-vous annuler cette commande?", vbYesNo) = vbYes Then
  Call AnnulerCommande
  End If
  End If
  End If
End If
End If

Exit Sub

Oups:

wOups "frmAchat", "lvwAchat_DblClick", Err, Err.number, Err.Description
End Sub

Private Sub lvwAchat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
 On Error GoTo Oups

 Dim iNbreSelected As Integer
 Dim iIndexSelected As Integer
 Dim iCompteur As Integer
 Dim bAfficherMenu As Boolean

 If m_eMode = MODE_AJOUT_MODIF Then
 If Button = vbRightButton Then
 If lvwAchat.ListItems.count > 0 Then
 'S'il y a plusieurs items de sélectionnés, c'est parce que l'utilisateur
 'a sélectionné plusieurs items
 'Donc, on ne désélectionne pas
 For iCompteur = 1 To lvwAchat.ListItems.count
 If lvwAchat.ListItems(iCompteur).Selected = True Then
 iNbreSelected = iNbreSelected + 1

  iIndexSelected = iCompteur
  End If
  Next

  If iNbreSelected = 1 Then
  lvwAchat.ListItems(iIndexSelected).Selected = False
  End If

  Set lvwAchat.DropHighlight = lvwAchat.HitTest(X, Y)

  If Not lvwAchat.DropHighlight Is Nothing Then
 If iNbreSelected = 1 Then
 lvwAchat.DropHighlight.Selected = True

 If lvwAchat.DropHighlight.SubItems(I_COL_ACHAT_DATE_REQUISE) = "" Then
 lvwAchat.DropHighlight.SubItems(I_COL_ACHAT_DATE_REQUISE) = " "
 End If

 If lvwAchat.DropHighlight.ListSubItems(I_COL_ACHAT_DATE_REQUISE).ForeColor = COLOR_ORANGE Then
 bAfficherMenu = True
 Else
 bAfficherMenu = False
 End If
 Else
 bAfficherMenu = False
 End If
 Else
 bAfficherMenu = False
 End If

 If bAfficherMenu = True Then
 Call RemplirOptionsMenuRightClick(iNbreSelected)

 Call PopupMenu(mnuRightClick)
1  End If
 End If
 Else
 If Shift <> vbCtrlMask And Shift <> vbShiftMask Then
 Set lvwAchat.DropHighlight = Nothing
 End If
 End If
End If

Exit Sub

Oups:

wOups "frmAchat", "lvwAchat_MouseDown", Err, Err.number, Err.Description
End Sub

Private Sub RemplirOptionsMenuRightClick(ByVal iNbreSelected As Integer)

 On Error GoTo Oups

 Dim bDateRequise As Boolean

 If iNbreSelected = 1 Then
 'Si c'est une sous-section
 Select Case lvwAchat.SelectedItem.ListSubItems(I_COL_ACHAT_PIECE).ForeColor
 Case COLOR_ORANGE:
 bDateRequise = True
 End Select
 End If

 'Pour empeche que tous les éléments deviennent invisible, je les mets visible au
 'début
 mnuDateRequise.Visible = True
 
 mnuDateRequise.Visible = bDateRequise

 Exit Sub

Oups:

 wOups "frmAchat", "RemplirOptionsMenuRightClick", Err, Err.number, Err.Description
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

  wOups "frmAchat", "lvwPieces_ColumnClick", Err, Err.number, Err.Description
End Sub

Private Sub RechercherPiece(ByVal iCol As Integer, ByVal sTexte As String)

 On Error GoTo Oups

 Dim sValeur As String
 Dim rstcat As ADODB.Recordset
 Dim iCompteur As Integer
 Dim bTrouverLvw As Boolean
 Dim bTrouverRst As Boolean
 Dim iIndexCat As Integer
 Dim sChamps As String
 Dim sCategorie As String
 
 For iCompteur = 1 To lvwPieces.ListItems.count
 If iCol > 0 Then
  sValeur = lvwPieces.ListItems(iCompteur).SubItems(iCol)
  Else
  sValeur = lvwPieces.ListItems(iCompteur).Text
  End If
 
  sValeur = UCase(sValeur)
  sTexte = UCase(sTexte)
 
  If InStr(1, sValeur, sTexte) > 0 Then
  lvwPieces.ListItems(iCompteur).Selected = True
 
 Call lvwPieces.SelectedItem.EnsureVisible
 
bTrouverLvw = True
 End If
 
 If bTrouverLvw = True Then
 Exit For
 End If
Next
 
Select Case iCol
 Case I_COL_PIECES_PIECE_GRB: sChamps = "PIECE_GRB"
 Case I_COL_PIECES_NO_ITEM: sChamps = "PIECE"
 Case I_COL_PIECES_MANUFACT: sChamps = "FABRICANT"
 Case I_COL_PIECES_DESCR_FR: sChamps = "DESC_FR"
 Case I_COL_PIECES_DESCR_EN: sChamps = "DESC_EN"
1  End Select
 
iIndexCat = cmbCategorie.ListIndex
 
 If bTrouverLvw = False Then
 Set rstcat = New ADODB.Recordset

 For iCompteur = iIndexCat + 1 To cmbCategorie.ListCount - 1
 sCategorie = Replace(cmbCategorie.LIST(iCompteur), "'", "''")
 
 If m_eCatalogue = ELECTRIQUE Then
1  Call rstcat.Open("SELECT * FROM GrbCatalogueElec WHERE INSTR(1, " & sChamps & ", '" & sTexte & "') > 0 AND CATEGORIE = '" & sCategorie & "'", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstcat.Open("SELECT * FROM GrbCatalogueMec WHERE INSTR(1, " & sChamps & ", '" & sTexte & "') > 0 AND CATEGORIE = '" & sCategorie & "'", g_connData, adOpenDynamic, adLockOptimistic)
 End If
 
 If Not rstcat.EOF Then
 bTrouverRst = True
 
 cmbCategorie.ListIndex = iCompteur
 
 Call RechercherPiece(iCol, sTexte)
 
 Exit For
 End If
 
 Call rstcat.Close
 Next
 
 If bTrouverRst = False Then
 For iCompteur = 0 To iIndexCat - 1
 sCategorie = Replace(cmbCategorie.LIST(iCompteur), "'", "''")
 
 If m_eCatalogue = ELECTRIQUE Then
 Call rstcat.Open("SELECT * FROM GrbCatalogueElec WHERE INSTR(1, " & sChamps & ", '" & sTexte & "') > 0 AND CATEGORIE = '" & sCategorie & "'", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstcat.Open("SELECT * FROM GrbCatalogueMec WHERE INSTR(1, " & sChamps & ", '" & sTexte & "') > 0 AND CATEGORIE = '" & sCategorie & "'", g_connData, adOpenDynamic, adLockOptimistic)
 End If
 
 If Not rstcat.EOF Then
 bTrouverRst = True
 
 cmbCategorie.ListIndex = iCompteur
 
 Call RechercherPiece(iCol, sTexte)
 
 Exit For
 End If
 
 Call rstcat.Close
 Next
 
 If bTrouverRst = False Then
 Call MsgBox("Aucun enregistrements trouvés!", vbOKOnly, "Erreur")
 End If
 End If

 Set rstcat = Nothing
3  End If

Exit Sub

Oups:

3  wOups "frmAchat", "RechercherPiece", Err, Err.number, Err.Description
End Sub

Private Sub lvwPieces_KeyDown(KeyCode As Integer, Shift As Integer)

 On Error GoTo Oups

 Dim sTexte As String
 
 If Shift = vbCtrlMask Then
 If KeyCode = vbKeyF Then
 sTexte = InputBox("Quel est le texte à rechercher?")
 
 If Trim$(sTexte) <> vbNullString Then
 If Len(Trim$(sTexte)) >= 3 Then
 Call RechercherPiece(I_COL_PIECES_NO_ITEM, sTexte)
 Else
 Call MsgBox("Il faut un minimum de 3 caractères pour rechercher!", vbOKOnly, "Erreur")
 End If
  End If
  End If
  End If

  Exit Sub

Oups:

  wOups "frmAchat", "lvwPieces_KeyDown", Err, Err.number, Err.Description
End Sub

Private Sub cmdEnregistrer_Click()

 On Error GoTo Oups

 Dim objControl As Control
 Dim iIndexAchat As Integer
 
 'Vérification des textbox
 Screen.MousePointer = vbHourglass
 
 For Each objControl In Me
 If TypeOf objControl Is TextBox Then
 If objControl.Visible = True Then
 If Trim$(objControl.Text) = vbNullString Then
 Screen.MousePointer = vbDefault

 Call MsgBox("Vous devez remplir tous les champs!", vbOKOnly, "Erreur")
 
 Exit Sub
  End If
  End If
  End If
  Next
 
  If BackupPieces(txtNoAchat.Text) = False Then
  Screen.MousePointer = vbDefault

  If MsgBox("Une erreur est survenue lors de la copie de sauvegarde de l'achat en cours!" & vbNewLine & _
 vbNewLine & _
 "Voulez-vous continuer?", vbYesNo) = vbNo Then
  Exit Sub
Else
Screen.MousePointer = vbHourglass
 End If
End If
 
 'Enregistre l'achat
Call EnregistrerAchat(txtNoAchat.Text)
 
 'Initialisation des variables booléennes
m_bInventaire = False
m_bMauvaisPrix = False
m_bPieceInutile = False
m_bRecherchePiece = False
 
Call OuvrirAchat(False)
 
 'Remet en mode inactif
Call AfficherControles(MODE_INACTIF)
 
 'Affiche l'achat actuel
If Len(txtNoAchat.Text) =   Then
iIndexAchat = TrouverNouvelIndex
End If

 If iIndexAchat > 0 Then
 Call AfficherAchat(txtNoAchat.Text & "-" & Right$("00" & iIndexAchat, 3))
 Else
 Call AfficherAchat(txtNoAchat.Text)
 End If
 
1  Screen.MousePointer = vbDefault

 Exit Sub

Oups:

 wOups "frmAchat", "cmdEnregistrer_Click", Err, Err.number, Err.Description
End Sub

Private Function BackupPieces(ByVal sNoAchat As String) As Boolean

 On Error GoTo Oups

 Dim rstAchat As ADODB.Recordset
 Dim rstAchatBackup As ADODB.Recordset
 Dim sDateCopie As String
 Dim sIDAchat As String
 Dim iIndexAchat As Integer

 If m_bModeAjout = False Then
 sIDAchat = Left$(sNoAchat, 9)

 iIndexAchat = Right$(sNoAchat, 3)
 Else
 BackupPieces = True

  Exit Function
  End If

  Set rstAchat = New ADODB.Recordset
  Set rstAchatBackup = New ADODB.Recordset

  Call rstAchat.Open("SELECT * FROM GrbAchat_Pieces WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat, g_connData, adOpenForwardOnly, adLockReadOnly)

  Call rstAchatBackup.Open("SELECT * FROM GrbAchat_Pieces_Tampon", g_connData, adOpenDynamic, adLockOptimistic)

  sDateCopie = ConvertDate(Date) & " " & Time

  Do While Not rstAchat.EOF
Call rstAchatBackup.AddNew

1 rstAchatBackup.Fields("DateCopie") = sDateCopie

 rstAchatBackup.Fields("IDAchat") = rstAchat.Fields("IDAchat")
 rstAchatBackup.Fields("IndexAchat") = rstAchat.Fields("IndexAchat")

 rstAchatBackup.Fields("Initiales") = g_sInitiale
 rstAchatBackup.Fields("PIECE") = rstAchat.Fields("PIECE")
 rstAchatBackup.Fields("NuméroLigne") = rstAchat.Fields("NuméroLigne")
 rstAchatBackup.Fields("Qté") = rstAchat.Fields("Qté")
 rstAchatBackup.Fields("Desc_FR") = rstAchat.Fields("Desc_FR")
 rstAchatBackup.Fields("Desc_EN") = rstAchat.Fields("Desc_EN")
 rstAchatBackup.Fields("Manufact") = rstAchat.Fields("Manufact")
 rstAchatBackup.Fields("Prix_list") = rstAchat.Fields("Prix_list")
rstAchatBackup.Fields("Escompte") = rstAchat.Fields("Escompte")
 rstAchatBackup.Fields("Prix_net") = rstAchat.Fields("Prix_net")
 rstAchatBackup.Fields("IDFRS") = rstAchat.Fields("IDFRS")
 rstAchatBackup.Fields("Prix_total") = rstAchat.Fields("Prix_total")
 rstAchatBackup.Fields("Type") = rstAchat.Fields("Type")
 rstAchatBackup.Fields("Commandé") = rstAchat.Fields("Commandé")
 rstAchatBackup.Fields("Retour") = rstAchat.Fields("Retour")
1  rstAchatBackup.Fields("NoRetour") = rstAchat.Fields("NoRetour")
 rstAchatBackup.Fields("Recu") = rstAchat.Fields("Recu")
 rstAchatBackup.Fields("DateRéception") = rstAchat.Fields("DateRéception")
 rstAchatBackup.Fields("QuantitéRecue") = rstAchat.Fields("QuantitéRecue")
 rstAchatBackup.Fields("DateCommande") = rstAchat.Fields("DateCommande")
 rstAchatBackup.Fields("DateRequise") = rstAchat.Fields("DateRequise")
 rstAchatBackup.Fields("Inutile") = rstAchat.Fields("Inutile")
 rstAchatBackup.Fields("CommandeAnnulée") = rstAchat.Fields("CommandeAnnulée")
 rstAchatBackup.Fields("DateRetour") = rstAchat.Fields("DateRetour")
 rstAchatBackup.Fields("PrixOrigine") = rstAchat.Fields("PrixOrigine")
 rstAchatBackup.Fields("Devise") = rstAchat.Fields("Devise")

 Call rstAchatBackup.Update

 Call rstAchat.MoveNext
2  Loop

Call rstAchat.Close
2  Set rstAchat = Nothing

Call rstAchatBackup.Close
2  Set rstAchatBackup = Nothing

BackupPieces = True

2  Exit Function

Oups:

wOups "frmAchat", "BackupPieces", Err, Err.number, Err.Description
End Function

Private Function TrouverNouvelIndex() As Integer

 On Error GoTo Oups

 Dim rstMax As ADODB.Recordset
 Dim iIndex As Integer
 
 Set rstMax = New ADODB.Recordset
 
 Call rstMax.Open("SELECT MAX(IndexAchat) AS MaxIndex FROM GrbAchat WHERE IDAchat = '" & txtNoAchat.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 iIndex = rstMax.Fields("MaxIndex")
 
 Call rstMax.Close
 Set rstMax = Nothing
 
 TrouverNouvelIndex = iIndex

 Exit Function

Oups:

 wOups "frmAchat", "TrouverNouvelIndex", Err, Err.number, Err.Description
End Function

Private Sub EnregistrerAchat(ByVal sNoAchat As String)

 On Error GoTo Oups

 Dim rstAchat As ADODB.Recordset
 Dim rstPiece As ADODB.Recordset
 Dim rstMax As ADODB.Recordset
 Dim itmPiece As ListItem
 Dim dblPrixTotal As Double
 Dim sIDAchat As String
 Dim iIndexAchat As Integer
 Dim iCompteur As Integer
 Dim testgll As String
 sIDAchat = Left$(sNoAchat, 9)
 
 Set rstAchat = New ADODB.Recordset
 
 'Si c'est un ajout
  If m_bModeAjout = True Then
 'On ouvre le recordset
  Call rstAchat.Open("SELECT * FROM GrbAchat", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Ajouter une nouvelle achat
  Call rstAchat.AddNew
 
  m_bModeAjout = False
  Else
  iIndexAchat = Right$(sNoAchat, 3)
 
  Call rstAchat.Open("SELECT * FROM GrbAchat WHERE IDAchat" & " = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat, g_connData, adOpenDynamic, adLockOptimistic)
 
 'Si c'est une modification, il faut effacer les pieces et remplir les nouvelles
  Call g_connData.Execute("DELETE * FROM GrbAchat_Pieces WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat)
10 End If
 
 'Enregistrement de l'achat
 
 'IDAchat
rstAchat.Fields("IDAchat") = sIDAchat
 
 'IndexAchat
If iIndexAchat = 0 Then
 Set rstMax = New ADODB.Recordset

 'Pour avoir le dernier index
 Call rstMax.Open("SELECT MAX(IndexAchat) As MaxAchat FROM GrbAchat WHERE IDAchat = '" & sNoAchat & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 If Not IsNull(rstMax("MaxAchat")) Then
 rstAchat.Fields("IndexAchat") = rstMax("MaxAchat") + 1
 Else
 rstAchat.Fields("IndexAchat") = 1
 End If
 
 Call rstMax.Close
 Set rstMax = Nothing
1  Else
 rstAchat.Fields("IndexAchat") = iIndexAchat
 End If
 
rstAchat.Fields("Raison") = txtRaison.Text
 rstAchat.Fields("DateAchat") = txtDate.Text
rstAchat.Fields("Acheteur") = txtAcheteur.Tag
 
 If m_eCatalogue = ELECTRIQUE Then
1  rstAchat.Fields("Type") = "E"
 Else
 rstAchat.Fields("Type") = "M"
End If

Set rstPiece = New ADODB.Recordset

Call rstPiece.Open("SELECT * FROM GrbAchat_Pieces", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Enregistrement des pièces
For iCompteur = 1 To lvwAchat.ListItems.count
 Set itmPiece = lvwAchat.ListItems(iCompteur)
 
 Call rstPiece.AddNew
 
 rstPiece.Fields("IDAchat") = rstAchat.Fields("IDAchat")
 rstPiece.Fields("IndexAchat") = rstAchat.Fields("IndexAchat")
 
 
 rstPiece.Fields("PIECE") = itmPiece.SubItems(I_COL_ACHAT_PIECE)
 rstPiece.Fields("NuméroLigne") = iCompteur
rstPiece.Fields("Qté") = itmPiece.Text
 rstPiece.Fields("Desc_FR") = itmPiece.SubItems(I_COL_ACHAT_DESCR)
rstPiece.Fields("Desc_EN") = itmPiece.ListSubItems(I_COL_ACHAT_DESCR).Tag
 rstPiece.Fields("Manufact") = itmPiece.SubItems(I_COL_ACHAT_MANUFACT)
rstPiece.Fields("Prix_list") = Conversion(itmPiece.SubItems(I_COL_ACHAT_PRIX_LIST), MODE_PAS_FORMAT, 4)

 If itmPiece.SubItems(I_COL_ACHAT_PRIX_LIST) <> vbNullString Then
 If itmPiece.ListSubItems(I_COL_ACHAT_PRIX_LIST).Tag <> vbNullString Then
 rstPiece.Fields("PrixOrigine") = Replace(Round(CDbl(Replace(itmPiece.ListSubItems(I_COL_ACHAT_PRIX_LIST).Tag, ".", ",")), 4), ".", ",")
 Else
 rstPiece.Fields("PrixOrigine") = "0"
 End If
 Else
 rstPiece.Fields("PrixOrigine") = "0"
 End If
 
 If itmPiece.SubItems(I_COL_ACHAT_TOTAL) <> "" Then
 rstPiece.Fields("Devise") = itmPiece.ListSubItems(I_COL_ACHAT_TOTAL).Tag
 Else
 rstPiece.Fields("Devise") = ""
 End If

 If Trim$(itmPiece.SubItems(I_COL_ACHAT_ESCOMPTE)) <> "" Then
 rstPiece.Fields("Escompte") = Conversion(Replace(itmPiece.SubItems(I_COL_ACHAT_ESCOMPTE), "%", "") / 100, MODE_PAS_FORMAT)
 Else
 rstPiece.Fields("Escompte") = ""
 End If

rstPiece.Fields("Prix_net") = Conversion(itmPiece.SubItems(I_COL_ACHAT_PRIX_NET), MODE_PAS_FORMAT, 4)
 rstPiece.Fields("DateRéception") = itmPiece.Tag
 rstPiece.Fields("NoRetour") = itmPiece.ListSubItems(I_COL_ACHAT_MANUFACT).Tag

 If itmPiece.ListSubItems(I_COL_ACHAT_DISTRIB).Tag <> "" Then
 rstPiece.Fields("IDFRS") = itmPiece.ListSubItems(I_COL_ACHAT_DISTRIB).Tag
4 Else
4 rstPiece.Fields("IDFRS") = 0
4 End If

4 If itmPiece.ListSubItems(I_COL_ACHAT_PIECE).ForeColor = COLOR_ORANGE Then
4 rstPiece.Fields("Commandé") = True
4 Else
4 rstPiece.Fields("Commandé") = False
4 End If

4 If itmPiece.ListSubItems(I_COL_ACHAT_PIECE).ForeColor = COLOR_BLEU Or itmPiece.ListSubItems(I_COL_ACHAT_PIECE).ForeColor = COLOR_GRIS Then
4 rstPiece.Fields("Recu") = True
4 Else
4  rstPiece.Fields("Recu") = False
4  End If

4  If itmPiece.ListSubItems(I_COL_ACHAT_PIECE).ForeColor = COLOR_ROUGE Then
4  rstPiece.Fields("Retour") = True
4  Else
4  rstPiece.Fields("Retour") = False
4  End If

4  If itmPiece.ListSubItems(I_COL_ACHAT_PIECE).ForeColor = COLOR_BRUN Then
50 rstPiece.Fields("Inutile") = True
5 Else
 rstPiece.Fields("Inutile") = False
 End If

 If itmPiece.ListSubItems(I_COL_ACHAT_PIECE).ForeColor = COLOR_VERT_FORET Then
 rstPiece.Fields("CommandeAnnulée") = True
 Else
 rstPiece.Fields("CommandeAnnulée") = False
 End If

 rstPiece.Fields("Prix_Total") = Conversion(itmPiece.SubItems(I_COL_ACHAT_TOTAL), MODE_PAS_FORMAT)

 If itmPiece.SubItems(I_COL_ACHAT_DATE_COMMANDE) <> "" Then
 rstPiece.Fields("DateCommande") = itmPiece.SubItems(I_COL_ACHAT_DATE_COMMANDE)
5  Else
5  rstPiece.Fields("DateCommande") = ""
5  End If
 
5  If itmPiece.SubItems(I_COL_ACHAT_DATE_REQUISE) <> "" Then
5  rstPiece.Fields("DateRequise") = itmPiece.SubItems(I_COL_ACHAT_DATE_REQUISE)
5  Else
5  rstPiece.Fields("DateRequise") = ""
5  End If
 
60 Call rstPiece.Update
 
  If rstPiece.Fields("Prix_Total") <> vbNullString Then
  dblPrixTotal = dblPrixTotal + rstPiece.Fields("Prix_Total")
  End If
  Next

  rstAchat.Fields("PrixTotal") = CStr(dblPrixTotal)
 
  Call rstAchat.Update
 
  Call rstAchat.Close
  Set rstAchat = Nothing
 
  Call rstPiece.Close
  Set rstPiece = Nothing

  Exit Sub

Oups:

6  wOups "frmAchat", "EnregistrerAchat", Err, Err.number, Err.Description

6  If Erl >= 230 And Erl <= 615 Then
6  Call MsgBox("La pièce " & itmPiece.SubItems(I_COL_ACHAT_PIECE) & " risque de contenir des erreurs." & vbNewLine & _
 "Il se peut qu'elle ne soit plus présente dans la liste.")
6  End If
 
6  Resume Next
End Sub

Private Sub AfficherAchat(ByVal sNoAchat As String)

 On Error GoTo Oups

 'Remet en mode affichage le projet ou l'achat voulue
 m_bModeAffichage = True
 
 'Vide les champs
 Call ViderChamps
 
 'Rempli le combo
 Call RemplirComboAchat(sNoAchat)

 Exit Sub

Oups:

 wOups "frmAchat", "AfficherAchat", Err, Err.number, Err.Description
End Sub

Private Sub Cmdfermer_Click()

 On Error GoTo Oups

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmAchat", "cmdFermer_Click", Err, Err.number, Err.Description
End Sub

Private Sub Cmdajouter_Click()

 On Error GoTo Oups

 'Ajoute une achat
 Dim rstEmploye As ADODB.Recordset
 
 'Initialisation des variables booléennes
 m_bInventaire = False
 m_bMauvaisPrix = False
 m_bPieceInutile = False
 m_bRecherchePiece = False
 
 Call frmAjoutAchat.Afficher(m_eCatalogue)
 
 If m_bAnnuler = False Then
 If m_sNoAchat <> vbNullString Then
 'Vide les champs
 Call ViderChamps
 
 txtAcheteur.Text = g_sEmploye

  Set rstEmploye = New ADODB.Recordset
 
  Call rstEmploye.Open("SELECT NoEmploye FROM GrbEmployés WHERE Initiale = '" & g_sInitiale & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
  txtAcheteur.Tag = rstEmploye.Fields("NoEmploye")
 
  Call rstEmploye.Close
  Set rstEmploye = Nothing
 
  txtDate.Text = ConvertDate(Date)
 
 'Pour pouvoir réafficher l'élément affiché avant l'ajout au cas où la personne
 'annule l'ajout
  m_sAncienAchat = txtNoAchat.Text
 
 'Affiche le nouveau numéro
  txtNoAchat.Text = m_sNoAchat
 
 m_bModeAjout = True
m_bModeAffichage = False
 
 'Met le form en mode ajout/modif
 Call AfficherControles(MODE_AJOUT_MODIF)
 End If
End If

Exit Sub

Oups:

wOups "frmAchat", "cmdAjouter_Click", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboCategorie()

 On Error GoTo Oups

 'Remplir le combo des tables (Pièces)
 Dim rstCategorie As ADODB.Recordset
 Dim sNomTable As String
 
 'Il faut vider le combo avant de le remplir
 Call cmbCategorie.Clear
 
 Set rstCategorie = New ADODB.Recordset

 'On rempli le recordset avec le nom de chaque catégorie
 If m_eCatalogue = ELECTRIQUE Then
 Call rstCategorie.Open("SELECT DISTINCT CATEGORIE FROM GrbCatalogueElec", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstCategorie.Open("SELECT DISTINCT CATEGORIE FROM GrbCatalogueMec", g_connData, adOpenDynamic, adLockOptimistic)
 End If
 
 'Tant que ce n'est pas la fin des enregistrements
 Do While Not rstCategorie.EOF
  If Not IsNull(rstCategorie.Fields("CATEGORIE")) Then
 'On ajoute le nom de la catégorie dans le combo
  Call cmbCategorie.AddItem(rstCategorie.Fields("CATEGORIE"))
  End If
 
  Call rstCategorie.MoveNext
  Loop
 
  Call rstCategorie.Close
  Set rstCategorie = Nothing

  If cmbCategorie.ListCount > 0 Then
cmbCategorie.ListIndex = 0
End If

Exit Sub

Oups:

wOups "frmAchat", "RemplirComboCategorie", Err, Err.number, Err.Description
End Sub

Private Sub cmdModifier_Click()

 On Error GoTo Oups

 Dim rstAchat As ADODB.Recordset
 Dim sIDAchat As String
 Dim iIndexAchat As Integer

 'Modifier un achat
 
 'Initialisation des variables booléennes
 m_bInventaire = False
 m_bMauvaisPrix = False
 m_bPieceInutile = False
 m_bRecherchePiece = False
 
 If cmbNoAchat.ListIndex > -1 Then
 sIDAchat = Trim$(Left$(txtNoAchat.Text, InStr(1, txtNoAchat.Text, "-") + 2))
 iIndexAchat = CInt(Trim$(Right$(txtNoAchat.Text, Len(txtNoAchat.Text) - InStrRev(txtNoAchat.Text, "-"))))

  Set rstAchat = New ADODB.Recordset

  Call rstAchat.Open("SELECT Modification, Par FROM GrbAchat WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat, g_connData, adOpenDynamic, adLockOptimistic)

  If rstAchat.Fields("Modification") = False Then
  Call rstAchat.Close
  Set rstAchat = Nothing

  Screen.MousePointer = vbHourglass
 
 'Pour pouvoir afficher le dernier enregistrement affiché quand la personne va
 'enregistrer ou annuler
  m_sAncienAchat = txtNoAchat.Text
 
  m_bModeAjout = False
 m_bModeAffichage = False

Call OuvrirAchat(True)
 
 Call AfficherControles(MODE_AJOUT_MODIF)
 
 Screen.MousePointer = vbDefault
 Else
 Call MsgBox("Cet achat est en modification par " & rstAchat.Fields("Par") & "!", vbOKOnly, "Erreur")

 Call rstAchat.Close
 Set rstAchat = Nothing
 End If
End If

Exit Sub

Oups:

wOups "frmAchat", "cmdModifier_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdsupprimer_Click()

 On Error GoTo Oups

 Dim iReponse As Integer
 Dim sIDAchat As String
 Dim iIndexAchat As Integer
 Dim rstAchat As ADODB.Recordset
 
 If cmbNoAchat.ListCount > 0 Then
 sIDAchat = Trim$(Left$(txtNoAchat.Text, InStr(1, txtNoAchat.Text, "-") + 2))
 iIndexAchat = CInt(Trim$(Right$(txtNoAchat.Text, Len(txtNoAchat.Text) - InStrRev(txtNoAchat.Text, "-"))))

 Set rstAchat = New ADODB.Recordset

 Call rstAchat.Open("SELECT Modification, Par FROM GrbAchat WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat, g_connData, adOpenDynamic, adLockOptimistic)

 If rstAchat.Fields("Modification") = False Then
  Call rstAchat.Close
  Set rstAchat = Nothing

 'Valider le choix
  iReponse = MsgBox("Voulez-vous vraiment effacer l'achat " & txtNoAchat.Text & "?", vbYesNo)
 
 'Si il veut vraiment effacer
  If iReponse = vbYes Then
 'Efface les pièces
  sIDAchat = Left$(txtNoAchat.Text, 9)
 
  iIndexAchat = Right$(txtNoAchat.Text, 3)
 
 'Efface les pièces
  Call g_connData.Execute("DELETE * FROM GrbAchat_Pieces WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat)
 
 'Efface l'achat
  Call g_connData.Execute("DELETE * FROM GrbAchat WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat)
 
 'Affiche la premiere achat
 Call RemplirComboAchat(vbNullString)
 End If
 Else
 Call MsgBox("Cet achat est en modification par " & rstAchat.Fields("Par") & "!", vbOKOnly, "Erreur")

 Call rstAchat.Close
 Set rstAchat = Nothing
 End If
End If

Exit Sub

Oups:

wOups "frmAchat", "cmdSupprimer_Click", Err, Err.number, Err.Description
End Sub

Private Sub ViderChamps()

 On Error GoTo Oups

 'Méthode qui initialise les champs
 txtPrixTotal.Text = 0
 txtDate.Text = vbNullString
 txtRaison.Text = vbNullString
 txtAcheteur.Text = vbNullString
 
 Call lvwAchat.ListItems.Clear

 Exit Sub

Oups:

 wOups "frmAchat", "ViderChamps", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboAchat(ByVal sNoAchat As String)

 On Error GoTo Oups

 'Rempli le combo des achats
 Dim rstAchat As ADODB.Recordset
 Dim sType As String
 Dim iCompteur As Integer
 
 'Il faut vider le combo avant de le remplir
 Call cmbNoAchat.Clear
 
 If m_eCatalogue = ELECTRIQUE Then
 sType = "E"
 Else
 sType = "M"
 End If
 
 Set rstAchat = New ADODB.Recordset
 
  Call rstAchat.Open("SELECT * FROM GrbAchat WHERE Type = '" & sType & "' ORDER BY IDAchat DESC, IndexAchat DESC", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Tant que ce n'est pas la fin des enregistrements
  Do While Not rstAchat.EOF
 'On met le numéro de l'achat dans le combo des achats
  Call cmbNoAchat.AddItem(rstAchat.Fields("IDAchat") & "-" & Right$("00" & rstAchat("IndexAchat"), 3))
 
  Call rstAchat.MoveNext
  Loop
 
  Call rstAchat.Close
  Set rstAchat = Nothing
 
 'Si le combo n'est pas vide, on sélectionne l'item voulu ou le premier
  If cmbNoAchat.ListCount > 0 Then
If sNoAchat <> vbNullString Then
For iCompteur = 0 To cmbNoAchat.ListCount - 1
 If cmbNoAchat.LIST(iCompteur) = sNoAchat Then
 cmbNoAchat.ListIndex = iCompteur

 Exit For
 End If
 Next
 Else
 cmbNoAchat.ListIndex = 0
 End If
Else
 Call ViderChamps
1  End If

Exit Sub

Oups:

 wOups "frmAchat", "RemplirComboAchat", Err, Err.number, Err.Description
End Sub

Private Sub CalculerPrixReel(ByVal sNoItem As String)

 On Error GoTo Oups

 Dim rstPieceFRS As ADODB.Recordset
 Dim rstConfig As ADODB.Recordset
 Dim sPrixCalcul As String
 Dim sTauxUSA As String
 Dim sTauxSPA As String
 Dim sType As String

 If m_eCatalogue = ELECTRIQUE Then
 sType = "E"
 Else
 sType = "M"
  End If

  Set rstConfig = New ADODB.Recordset

  Call rstConfig.Open("SELECT TauxAmericain, TauxEspagnol FROM GrbConfig", g_connData, adOpenForwardOnly, adLockReadOnly)

  sTauxUSA = rstConfig.Fields("TauxAmericain")
  sTauxSPA = rstConfig.Fields("TauxEspagnol")

  Call rstConfig.Close
  Set rstConfig = Nothing
 
  Set rstPieceFRS = New ADODB.Recordset

10 rstPieceFRS.CursorLocation = adUseServer
 
Call rstPieceFRS.Open("SELECT PrixReel, PRIX_NET, PRIX_SP, DeviseMonétaire FROM GrbPiecesFRS WHERE PIECE = '" & Replace(sNoItem, "'", "''") & "' AND Type = '" & sType & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
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
 rstPieceFRS.Fields("PrixReel") = Conversion(sPrixCalcul, MODE_DECIMAL, 4)
 End If
1  End If
 
 Call rstPieceFRS.Update
 
 Call rstPieceFRS.MoveNext
Loop
 
Call rstPieceFRS.Close
Set rstPieceFRS = Nothing

Exit Sub

Oups:

wOups "frmAchat", "CalculerPrixReel", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListViewFournisseur()

 On Error GoTo Oups

 'Rempli le listview des distributeur pour une pièce choisie
 Dim rstPieceFRS As ADODB.Recordset
 Dim rstContact As ADODB.Recordset
 Dim rstFRS As ADODB.Recordset
 Dim iCompteur As Integer
 Dim itmFRS As ListItem
 Dim sDevise As String
 Dim iNoClient As Integer
 Dim lColor As Long
 Dim sType As String
 
 If m_eCatalogue = ELECTRIQUE Then
  sType = "E"
  Else
  sType = "M"
  End If
 
 'vide le lister
  Call lvwfournisseur.ListItems.Clear
 
  Set rstPieceFRS = New ADODB.Recordset
 
  If m_bPieceInutile = True Then
  Call CalculerPrixReel(Trim$(lvwAchat.SelectedItem.SubItems(I_COL_ACHAT_PIECE)))

Call rstPieceFRS.Open("SELECT GrbPiecesFRS.*, GrbFournisseur.NomFournisseur FROM GrbPiecesFRS INNER JOIN GrbFournisseur ON GrbPiecesFRS.IDFRS = GrbFournisseur.IDFRS WHERE PIECE = '" & Trim$(Replace(lvwAchat.SelectedItem.SubItems(I_COL_ACHAT_PIECE), "'", "''")) & "' AND Type = '" & sType & "' ORDER BY PrixReel", g_connData, adOpenDynamic, adLockOptimistic)
Else
 If m_bInventaire = True Then
 Call CalculerPrixReel(Trim$(lvwInventaire.SelectedItem.Text))

 Call rstPieceFRS.Open("SELECT GrbPiecesFRS.*, GrbFournisseur.NomFournisseur FROM GrbPiecesFRS INNER JOIN GrbFournisseur ON GrbPiecesFRS.IDFRS = GrbFournisseur.IDFRS WHERE PIECE = '" & Trim$(Replace(lvwInventaire.SelectedItem.Text, "'", "''")) & "' AND Type = '" & sType & "' ORDER BY PrixReel", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 If m_bRecherchePiece = True Then
 Call CalculerPrixReel(Trim$(lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM)))

 Call rstPieceFRS.Open("SELECT GrbPiecesFRS.*, GrbFournisseur.NomFournisseur FROM GrbPiecesFRS INNER JOIN GrbFournisseur ON GrbPiecesFRS.IDFRS = GrbFournisseur.IDFRS WHERE PIECE = '" & Trim$(Replace(lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM), "'", "''")) & "' AND Type = '" & sType & "' ORDER BY PrixReel", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call CalculerPrixReel(Trim$(lvwPieces.SelectedItem.SubItems(I_COL_PIECES_NO_ITEM)))

 Call rstPieceFRS.Open("SELECT GrbPiecesFRS.*, GrbFournisseur.NomFournisseur FROM GrbPiecesFRS INNER JOIN GrbFournisseur ON GrbPiecesFRS.IDFRS = GrbFournisseur.IDFRS WHERE PIECE = '" & Trim$(Replace(lvwPieces.SelectedItem.SubItems(I_COL_PIECES_NO_ITEM), "'", "''")) & "' AND Type = '" & sType & "' ORDER BY PrixReel", g_connData, adOpenDynamic, adLockOptimistic)
 End If
 End If
 End If

Set rstFRS = New ADODB.Recordset

 Call rstFRS.Open("SELECT IDFRS FROM GrbFournisseur WHERE NomFournisseur = 'FOURNI PAR LE CLIENT'", g_connData, adOpenDynamic, adLockOptimistic)
 
iNoClient = rstFRS.Fields("IDFRS")

 Call rstFRS.Close
1  Set rstFRS = Nothing

 'tant il y a des fournisseur de la piece, ajoute dans lister
 Do While Not rstPieceFRS.EOF
 If rstPieceFRS.Fields("IDFRS") = iNoClient Then
 Call rstPieceFRS.MoveNext

 If rstPieceFRS.EOF Then
 Exit Do
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
 
 Set itmFRS = lvwfournisseur.ListItems.Add
 
itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = " "
3 itmFRS.SubItems(I_COL_FRS_ESCOMPTE) = " "
 itmFRS.SubItems(I_COL_FRS_PRIX_NET) = " "
 itmFRS.SubItems(I_COL_FRS_PRIX_SP) = " "
 
 'Nom du FRS
 itmFRS.Text = rstPieceFRS.Fields("NomFournisseur")
 
 itmFRS.Tag = rstPieceFRS.Fields("IDFRS")

 itmFRS.ForeColor = lColor
 
 'Personne ressource
 If Trim(rstPieceFRS.Fields("PERS_RESS")) <> vbNullString Then
 Set rstContact = New ADODB.Recordset

 Call rstContact.Open("SELECT NomContact FROM GrbContact WHERE IDContact = " & rstPieceFRS.Fields("PERS_RESS"), g_connData, adOpenDynamic, adLockOptimistic)
 
 If Not rstContact.EOF Then
 itmFRS.SubItems(I_COL_FRS_PERS_RESS) = rstContact.Fields("NomContact")

 itmFRS.ListSubItems(I_COL_FRS_PERS_RESS).ForeColor = lColor
 Else
 itmFRS.SubItems(I_COL_FRS_PERS_RESS) = ""
 End If
 
 Call rstContact.Close
 Set rstContact = Nothing
 End If
 
 'Date
 If Not IsNull(rstPieceFRS.Fields("Date")) Then
 itmFRS.SubItems(I_COL_FRS_DATE) = rstPieceFRS.Fields("Date")

4 itmFRS.ListSubItems(I_COL_FRS_DATE).ForeColor = lColor
4 Else
4 itmFRS.SubItems(I_COL_FRS_DATE) = vbNullString
4 End If
 
 'Entrer par
4 If Not IsNull(rstPieceFRS.Fields("Entrer_Par")) Then
4 itmFRS.SubItems(I_COL_FRS_ENTRER_PAR) = rstPieceFRS.Fields("Entrer_Par")

4 itmFRS.ListSubItems(I_COL_FRS_ENTRER_PAR).ForeColor = lColor
4 Else
4 itmFRS.SubItems(I_COL_FRS_ENTRER_PAR) = vbNullString
4 End If
  
 'Valide
4 If Not IsNull(rstPieceFRS.Fields("Valide")) Then
4  itmFRS.SubItems(I_COL_FRS_VALIDE) = rstPieceFRS.Fields("Valide")

4  itmFRS.ListSubItems(I_COL_FRS_VALIDE).ForeColor = lColor
4  Else
4  itmFRS.SubItems(I_COL_FRS_VALIDE) = vbNullString
4  End If
 
 'Prix listé
4  If rstPieceFRS.Fields("PRIX_LIST") <> vbNullString Then
4  itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = Conversion(Round(Replace(rstPieceFRS.Fields("PRIX_LIST"), ".", ","), 4), MODE_ARGENT, 4)

4  itmFRS.ListSubItems(I_COL_FRS_PRIX_LIST).ForeColor = lColor
50 End If
 
 'Escompte
5 If rstPieceFRS.Fields("ESCOMPTE") <> vbNullString Then
 itmFRS.SubItems(I_COL_FRS_ESCOMPTE) = Conversion(Replace(Replace(rstPieceFRS.Fields("ESCOMPTE"), "_", vbNullString), ".", ",") * 100, MODE_POURCENT)

 itmFRS.ListSubItems(I_COL_FRS_ESCOMPTE).ForeColor = lColor
 End If
 
 'Prix net
 If rstPieceFRS.Fields("PRIX_NET") <> vbNullString Then
 itmFRS.SubItems(I_COL_FRS_PRIX_NET) = Conversion(Round(Replace(rstPieceFRS.Fields("PRIX_NET"), ".", ","), 4), MODE_ARGENT, 4)

 itmFRS.ListSubItems(I_COL_FRS_PRIX_NET).ForeColor = lColor
 End If
 
 'Prix spécial
 If rstPieceFRS.Fields("PRIX_SP") <> vbNullString Then
 itmFRS.SubItems(I_COL_FRS_PRIX_SP) = Conversion(Round(rstPieceFRS.Fields("PRIX_SP"), 4), MODE_ARGENT, 4)

 itmFRS.ListSubItems(I_COL_FRS_PRIX_SP).ForeColor = lColor
5  End If

 'Quoter
5  If rstPieceFRS.Fields("QUOTER") = True Then
5  itmFRS.SubItems(I_COL_FRS_QUOTER) = "Oui"
5  Else
5  itmFRS.SubItems(I_COL_FRS_QUOTER) = "Non"
5  End If

5  itmFRS.ListSubItems(I_COL_FRS_QUOTER).ForeColor = lColor
 
 'Pour garder en mémoire le prix d'origine, je le mets dans le
 'tag de la colonne Prix Listé
5  If itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = vbNullString Then
60 itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = " "
  End If

  If rstPieceFRS.Fields("PRIX_NET") <> vbNullString Then
  If rstPieceFRS.Fields("PRIX_LIST") = "0,00" Or rstPieceFRS.Fields("PRIX_LIST") = "0" Then
  itmFRS.ListSubItems(I_COL_FRS_PRIX_LIST).Tag = Replace(rstPieceFRS.Fields("PRIX_NET"), ".", ",")
  Else
  itmFRS.ListSubItems(I_COL_FRS_PRIX_LIST).Tag = Replace(rstPieceFRS.Fields("PRIX_LIST"), ".", ",")
  End If
  Else
  itmFRS.ListSubItems(I_COL_FRS_PRIX_LIST).Tag = Replace(rstPieceFRS.Fields("PRIX_SP"), ".", ",")
  End If

  If itmFRS.SubItems(I_COL_FRS_PERS_RESS) = "" Then
6  itmFRS.SubItems(I_COL_FRS_PERS_RESS) = " "
6  End If

6  itmFRS.ListSubItems(I_COL_FRS_PERS_RESS).Tag = sDevise

6  Call rstPieceFRS.MoveNext
6  Loop
 
 'ferme la table
6  Call rstPieceFRS.Close
6  Set rstPieceFRS = Nothing

6  If m_bPieceInutile = False Then
70 Set itmFRS = lvwfournisseur.ListItems.Add

  itmFRS.Text = "CHOISIR ULTÉRIEUREMENT"

  itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = " "
  itmFRS.SubItems(I_COL_FRS_ESCOMPTE) = " "
  itmFRS.SubItems(I_COL_FRS_PRIX_NET) = " "
  itmFRS.SubItems(I_COL_FRS_PRIX_SP) = " "
  End If

  Exit Sub

Oups:

  wOups "frmAchat", "RemplirListViewFournisseur", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListViewInventaire()

 On Error GoTo Oups

 'Rempli le listview des pièces à commander dans l'inventaire
 Dim rstInv As ADODB.Recordset
 Dim itmInv As ListItem
 Dim lStock As Long
 Dim lMinimum As Long

 'Il faut vider le ListView avant de le remplir
 Call lvwInventaire.ListItems.Clear
 
 Set rstInv = New ADODB.Recordset
 
 If m_eCatalogue = ELECTRIQUE Then
 Call rstInv.Open("SELECT * FROM GrbInventaireElec WHERE Minimum = True ORDER BY NoItem", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstInv.Open("SELECT * FROM GrbInventaireMec WHERE Minimum = True ORDER BY NoItem", g_connData, adOpenDynamic, adLockOptimistic)
  End If
 
 'Tant que ce n'est pas la fin des enregistrements
  Do While Not rstInv.EOF
  If Not IsNull(rstInv.Fields("QuantitéStock")) Then
  lStock = Replace(rstInv.Fields("QuantitéStock"), ".", ",")
  Else
  lStock = 0
  End If
 
  If Not IsNull(rstInv.Fields("QuantitéMinimum")) Then
 lMinimum = rstInv.Fields("QuantitéMinimum")
1 Else
 lMinimum = 0
 End If
 
 If lStock < lMinimum Then
 'On l'ajoute
 Set itmInv = lvwInventaire.ListItems.Add
 
 'No piece
 If Not IsNull(rstInv.Fields("NoItem")) Then
 itmInv.Text = rstInv.Fields("NoItem")
 Else
 itmInv.Text = vbNullString
 End If
 
 'Fabricant
 If Not IsNull(rstInv.Fields("Manufacturier")) Then
 itmInv.SubItems(I_COL_INV_MANUFACT) = rstInv.Fields("Manufacturier")
 Else
 itmInv.SubItems(I_COL_INV_MANUFACT) = vbNullString
 End If
 
 'Description
 If Not IsNull(rstInv.Fields("Description")) Then
 itmInv.SubItems(I_COL_INV_DESCR) = rstInv.Fields("Description")
 Else
1  itmInv.SubItems(I_COL_INV_DESCR) = vbNullString
 End If

 'Commentaire
 If Not IsNull(rstInv.Fields("Commentaires")) Then
 itmInv.SubItems(I_COL_INV_COMMENT) = rstInv.Fields("Commentaires")
 Else
 itmInv.SubItems(I_COL_INV_COMMENT) = ""
 End If

 'Quantité en stock
 itmInv.SubItems(I_COL_INV_QTE_STOCK) = lStock

 'Quantité minimum
 itmInv.SubItems(I_COL_INV_QTE_MINIMUM) = lMinimum

 'Quantité à commander
 If Not IsNull(rstInv.Fields("Commande")) Then
 itmInv.SubItems(I_COL_INV_QTE_COMMANDE) = rstInv.Fields("Commande")
 Else
 itmInv.SubItems(I_COL_INV_QTE_COMMANDE) = vbNullString
 End If
 End If
 
Call rstInv.MoveNext
Loop
 
2  Call rstInv.Close
Set rstInv = Nothing

2  Exit Sub

Oups:

wOups "frmAchat", "RemplirListViewInventaire", Err, Err.number, Err.Description
End Sub


Private Sub RemplirListViewPieces()

 On Error GoTo Oups

 'Rempli le listview des pièces selon la catégorie de pièce choisit
 Dim rstPieces As ADODB.Recordset
 Dim itmPieces As ListItem
 Dim iIndex As Integer
 Dim bDebut As Boolean
 Dim sTri As String
 Dim sOrderBy As String
 Dim sCategorie As String
 
 sTri = m_sTri
 
 'Il faut vider le ListView avant de le remplir
 Call lvwPieces.ListItems.Clear

 Set rstPieces = New ADODB.Recordset
 
  Select Case cmbTri.ListIndex
 Case I_CMB_PIECE_GRB: sOrderBy = "PIECE_GRB"
  Case I_CMB_PIECE: sOrderBy = "PIECE"
  Case I_CMB_FABRICANT: sOrderBy = "FABRICANT"
  Case I_CMB_DESCR_FR: sOrderBy = "DESC_FR"
  Case I_CMB_DESCR_EN: sOrderBy = "DESC_EN"
  End Select
 
  sCategorie = Replace(cmbCategorie.Text, "'", "''")
 
  If m_eCatalogue = ELECTRIQUE Then
Call rstPieces.Open("SELECT * FROM GrbCatalogueElec WHERE CATEGORIE = '" & sCategorie & "' ORDER BY " & sOrderBy, g_connData, adOpenDynamic, adLockOptimistic)
Else
 Call rstPieces.Open("SELECT * FROM GrbCatalogueMec WHERE CATEGORIE = '" & sCategorie & "' ORDER BY " & sOrderBy, g_connData, adOpenDynamic, adLockOptimistic)
End If
 
iIndex = 1
 
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
1  bDebut = True
 End If
 
 'Si c'est la colonne No. d'item
 Case I_COL_PIECES_DESCR_FR:
 'Si la description française contient la recherche
 If InStr(1, UCase(rstPieces.Fields("DESC_FR")), UCase(sTri)) > 0 Then
 'On met la variable à true pour l'ajouter au début
 bDebut = True
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
 
 'Piece_GRB
 If Not IsNull(rstPieces.Fields("PIECE_GRB")) Then
 itmPieces.Text = rstPieces.Fields("PIECE_GRB")
Else
 itmPieces.Text = vbNullString
 End If
 
 'No piece
 If Not IsNull(rstPieces.Fields("PIECE")) Then
 itmPieces.SubItems(I_COL_PIECES_NO_ITEM) = rstPieces.Fields("PIECE")
 Else
 itmPieces.SubItems(I_COL_PIECES_NO_ITEM) = vbNullString
 End If
 
 'Fabricant
 If Not IsNull(rstPieces.Fields("FABRICANT")) Then
 itmPieces.SubItems(I_COL_PIECES_MANUFACT) = rstPieces.Fields("FABRICANT")
 Else
 itmPieces.SubItems(I_COL_PIECES_MANUFACT) = vbNullString
 End If
 
 'Description en francais
 If Not IsNull(rstPieces.Fields("DESC_FR")) Then
 itmPieces.SubItems(I_COL_PIECES_DESCR_FR) = rstPieces.Fields("DESC_FR")
 Else
 itmPieces.SubItems(I_COL_PIECES_DESCR_FR) = vbNullString
 End If
 
 'Description en anglais
 If Not IsNull(rstPieces.Fields("DESC_EN")) Then
 itmPieces.SubItems(I_COL_PIECES_DESCR_EN) = rstPieces.Fields("DESC_EN")
4 Else
4 itmPieces.SubItems(I_COL_PIECES_DESCR_EN) = vbNullString
4 End If
4 End If

 'Commentaire
4 If Not IsNull(rstPieces.Fields("COMMENTAIRE")) Then
4 itmPieces.SubItems(I_COL_PIECES_COMMENT) = rstPieces.Fields("COMMENTAIRE")
4 Else
4 itmPieces.SubItems(I_COL_PIECES_COMMENT) = ""
4 End If
 
4 Call rstPieces.MoveNext
4 Loop
 
4  Call rstPieces.Close
4  Set rstPieces = Nothing

4  Exit Sub

Oups:

4  wOups "frmAchat", "RemplirListViewPieces", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListViewAchat()

 On Error GoTo Oups

 'Remplis les pièces de l'achat avec la BD
 Dim rstAchat As ADODB.Recordset
 Dim rstFRS As ADODB.Recordset
 Dim itmAchat As ListItem
 Dim sIDAchat As String
 Dim iIndexAchat As Integer
 Dim lColor As Long
 Dim bBold As Boolean
 
 Call lvwAchat.ListItems.Clear
 
 sIDAchat = Left$(txtNoAchat.Text, 9)
 
 iIndexAchat = CInt(Right$(txtNoAchat.Text, 3))
 
  Set rstAchat = New ADODB.Recordset
 
  Call rstAchat.Open("SELECT * FROM GrbAchat_Pieces WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat & " ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
 
  Do While Not rstAchat.EOF
  bBold = False

  If rstAchat.Fields("Retour") = True Then
  lColor = COLOR_ROUGE
  Else
  If rstAchat.Fields("Recu") = True Then
 lColor = COLOR_GRIS 'Gris
Else
 If rstAchat.Fields("Inutile") = True Then
 lColor = COLOR_BRUN
 Else
 If rstAchat.Fields("IDFRS") = 0 Then
 lColor = COLOR_MAGENTA
 Else
 If rstAchat.Fields("Commandé") = True Then
 lColor = COLOR_ORANGE 'COLOR_ORANGE
 Else
 If rstAchat.Fields("CommandeAnnulée") = True Then
 lColor = COLOR_VERT_FORET
 bBold = True
 Else
 lColor = COLOR_NOIR
 End If
 End If
 End If
1  End If
 End If
 End If

 Set itmAchat = lvwAchat.ListItems.Add
 
 'Quantité
 If Not IsNull(rstAchat.Fields("Qté")) Then
 itmAchat.Text = rstAchat.Fields("Qté")
 Else
 itmAchat.Text = vbNullString
 End If

 itmAchat.ForeColor = lColor
 itmAchat.Bold = bBold
 
 itmAchat.Tag = rstAchat.Fields("DateRéception")
 
 'Numéro d'item
 If Not IsNull(rstAchat.Fields("PIECE")) Then
 itmAchat.SubItems(I_COL_ACHAT_PIECE) = rstAchat.Fields("PIECE")
 Else
 itmAchat.SubItems(I_COL_ACHAT_PIECE) = vbNullString
 End If

itmAchat.ListSubItems(I_COL_ACHAT_PIECE).ForeColor = lColor
 itmAchat.ListSubItems(I_COL_ACHAT_PIECE).Bold = bBold
 
 'Description en francais
If Not IsNull(rstAchat.Fields("DESC_FR")) Then
 itmAchat.SubItems(I_COL_ACHAT_DESCR) = rstAchat.Fields("DESC_FR")
Else
itmAchat.SubItems(I_COL_ACHAT_DESCR) = vbNullString
 End If

 itmAchat.ListSubItems(I_COL_ACHAT_DESCR).ForeColor = lColor
 itmAchat.ListSubItems(I_COL_ACHAT_DESCR).Bold = bBold
 
 'On met la description en anglais dans le tag de la description en francais
 If Not IsNull(rstAchat.Fields("Desc_EN")) Then
 itmAchat.ListSubItems(I_COL_ACHAT_DESCR).Tag = rstAchat.Fields("Desc_EN")
 Else
 itmAchat.ListSubItems(I_COL_ACHAT_DESCR).Tag = vbNullString
 End If
 
 'Fabricant
 If Not IsNull(rstAchat.Fields("Manufact")) Then
 itmAchat.SubItems(I_COL_ACHAT_MANUFACT) = rstAchat.Fields("Manufact")
Else
 itmAchat.SubItems(I_COL_ACHAT_MANUFACT) = vbNullString
End If

 itmAchat.ListSubItems(I_COL_ACHAT_MANUFACT).ForeColor = lColor
itmAchat.ListSubItems(I_COL_ACHAT_MANUFACT).Bold = bBold

 itmAchat.ListSubItems(I_COL_ACHAT_MANUFACT).Tag = rstAchat.Fields("NoRetour")
 
 'Prix listé
 If Trim(rstAchat.Fields("Prix_List")) <> vbNullString Then
 itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = Conversion(rstAchat.Fields("Prix_list"), MODE_ARGENT, 4)
Else
4 itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = " "
4 End If

4 itmAchat.ListSubItems(I_COL_ACHAT_PRIX_LIST).ForeColor = lColor

4 itmAchat.ListSubItems(I_COL_ACHAT_PRIX_LIST).Bold = bBold

4 itmAchat.ListSubItems(I_COL_ACHAT_PRIX_LIST).Tag = rstAchat.Fields("PrixOrigine")
 
 'Escompte
4 If Trim(rstAchat.Fields("Escompte")) <> vbNullString Then
4 itmAchat.SubItems(I_COL_ACHAT_ESCOMPTE) = Conversion(rstAchat.Fields("Escompte"), MODE_POURCENT)
4 Else
4 itmAchat.SubItems(I_COL_ACHAT_ESCOMPTE) = " "
4 End If

4 itmAchat.ListSubItems(I_COL_ACHAT_ESCOMPTE).ForeColor = lColor
4  itmAchat.ListSubItems(I_COL_ACHAT_ESCOMPTE).Bold = bBold
 
 'Prix net
4  If Trim(rstAchat.Fields("Prix_net")) <> vbNullString Then
4  itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(rstAchat.Fields("Prix_net"), MODE_ARGENT, 4)
4  Else
4  itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = " "
4  End If

4  itmAchat.ListSubItems(I_COL_ACHAT_PRIX_NET).ForeColor = lColor
4  itmAchat.ListSubItems(I_COL_ACHAT_PRIX_NET).Bold = bBold
 
 'Fournisseur
50 If Not IsNull(rstAchat.Fields("IDFRS")) Then
If rstAchat.Fields("IDFRS") <> 0 Then
 Set rstFRS = New ADODB.Recordset

 Call rstFRS.Open("SELECT NomFournisseur FROM GrbFournisseur WHERE IDFRS = " & rstAchat.Fields("IDFRS"), g_connData, adOpenDynamic, adLockOptimistic)
 
 'On affiche le nom dans la colonne
 If Not rstFRS.EOF Then
 itmAchat.SubItems(I_COL_ACHAT_DISTRIB) = rstFRS.Fields("NomFournisseur")
 Else
 itmAchat.SubItems(I_COL_ACHAT_DISTRIB) = ""
 End If
 
 'On affiche l'Id dans le tag
 itmAchat.ListSubItems(I_COL_ACHAT_DISTRIB).Tag = rstAchat.Fields("IDFRS")
 
 Call rstFRS.Close
 Set rstFRS = Nothing
5  Else
5  itmAchat.SubItems(I_COL_ACHAT_DISTRIB) = " "
5  End If
5  Else
5  itmAchat.SubItems(I_COL_ACHAT_DISTRIB) = vbNullString
5  End If

5  itmAchat.ListSubItems(I_COL_ACHAT_DISTRIB).ForeColor = lColor
5  itmAchat.ListSubItems(I_COL_ACHAT_DISTRIB).Bold = bBold
 
 'Prix total
60 If rstAchat.Fields("Prix_total") <> vbNullString Then
  itmAchat.SubItems(I_COL_ACHAT_TOTAL) = Conversion(Round(rstAchat.Fields("Prix_total"), 2), MODE_ARGENT)
  Else
  itmAchat.SubItems(I_COL_ACHAT_TOTAL) = " "
  End If

  itmAchat.ListSubItems(I_COL_ACHAT_TOTAL).ForeColor = lColor
  itmAchat.ListSubItems(I_COL_ACHAT_TOTAL).Bold = bBold

  itmAchat.ListSubItems(I_COL_ACHAT_TOTAL).Tag = rstAchat.Fields("Devise")

 'Date Commande
  If rstAchat.Fields("DateCommande") <> vbNullString Then
  itmAchat.SubItems(I_COL_ACHAT_DATE_COMMANDE) = rstAchat.Fields("DateCommande")
  Else
  itmAchat.SubItems(I_COL_ACHAT_DATE_COMMANDE) = ""
6  End If

6  itmAchat.ListSubItems(I_COL_ACHAT_DATE_COMMANDE).ForeColor = lColor
6  itmAchat.ListSubItems(I_COL_ACHAT_DATE_COMMANDE).Bold = bBold

 'Date Requise
6  If rstAchat.Fields("DateRequise") <> vbNullString Then
6  itmAchat.SubItems(I_COL_ACHAT_DATE_REQUISE) = rstAchat.Fields("DateRequise")
6  Else
6  itmAchat.SubItems(I_COL_ACHAT_DATE_REQUISE) = ""
6  End If

70 itmAchat.ListSubItems(I_COL_ACHAT_DATE_REQUISE).ForeColor = lColor
  itmAchat.ListSubItems(I_COL_ACHAT_DATE_REQUISE).Bold = bBold

  Call rstAchat.MoveNext
  Loop
 
  Call rstAchat.Close
  Set rstAchat = Nothing

  Exit Sub

Oups:

  wOups "frmAchat", "RemplirListViewAchat", Err, Err.number, Err.Description
End Sub

Private Sub lvwFournisseur_DblClick()

 On Error GoTo Oups

 If m_bPieceInutile = True Then
 Call ChoisirFournisseurMateriel
 Else
 Call ChoisirFournisseur
 End If

 Exit Sub

Oups:

 wOups "frmAchat", "lvwFournisseur_DblClick", Err, Err.number, Err.Description
End Sub

Private Sub ChoisirFournisseur()

 On Error GoTo Oups

 'On ajoute la pièce dans lvwAchat
 Dim rstConfig As ADODB.Recordset
 Dim sTauxUSA As String
 Dim sTauxSPA As String
 Dim sQuantite As String
 Dim itmAchat As ListItem
 Dim lColor As Long
 
 'Saisie de la quantité
 sQuantite = InputBox("Quelle est la quantité?", , m_sQuantite)

 sQuantite = Replace(sQuantite, ".", ",")

 Set rstConfig = New ADODB.Recordset

 Call rstConfig.Open("SELECT TauxAmericain, TauxEspagnol FROM GrbConfig", g_connData, adOpenForwardOnly, adLockReadOnly)

  sTauxUSA = rstConfig.Fields("TauxAmericain")
  sTauxSPA = rstConfig.Fields("TauxEspagnol")

  Call rstConfig.Close
  Set rstConfig = Nothing
 
  If sQuantite <> vbNullString Then
  If Not IsNumeric(sQuantite) Then
  Call MsgBox("Quantité non numérique!", vbOKOnly, "Erreur")
 
  Exit Sub
End If
Else
 Exit Sub
End If
 
Set itmAchat = lvwAchat.ListItems.Add
 
If lvwfournisseur.SelectedItem.Text = "CHOISIR ULTÉRIEUREMENT" Then
 lColor = COLOR_MAGENTA
Else
 lColor = COLOR_NOIR
End If
 
 'Quantité
itmAchat.Text = sQuantite
itmAchat.ForeColor = lColor
 
 'Numéro d'item
1  If m_bRecherchePiece = True Then
 itmAchat.SubItems(I_COL_ACHAT_PIECE) = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM)
 itmAchat.SubItems(I_COL_ACHAT_DESCR) = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_DESCR_FR)
 itmAchat.SubItems(I_COL_ACHAT_MANUFACT) = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_MANUFACT)
 itmAchat.ListSubItems(I_COL_ACHAT_DESCR).Tag = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_DESCR_EN)
Else
 If m_bInventaire = True Then
1  itmAchat.SubItems(I_COL_ACHAT_PIECE) = lvwInventaire.SelectedItem.Text
 itmAchat.SubItems(I_COL_ACHAT_DESCR) = lvwInventaire.SelectedItem.SubItems(I_COL_INV_DESCR)
 itmAchat.SubItems(I_COL_ACHAT_MANUFACT) = lvwInventaire.SelectedItem.SubItems(I_COL_INV_MANUFACT)
 Else
 itmAchat.SubItems(I_COL_ACHAT_PIECE) = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_NO_ITEM)
 itmAchat.SubItems(I_COL_ACHAT_DESCR) = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_DESCR_FR)
 itmAchat.SubItems(I_COL_ACHAT_MANUFACT) = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_MANUFACT)
 itmAchat.ListSubItems(I_COL_ACHAT_DESCR).Tag = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_DESCR_EN)
 End If
End If

itmAchat.ListSubItems(I_COL_ACHAT_PIECE).ForeColor = lColor
itmAchat.ListSubItems(I_COL_ACHAT_DESCR).ForeColor = lColor
itmAchat.ListSubItems(I_COL_ACHAT_MANUFACT).ForeColor = lColor
 
 'Prix listé
2  If Trim$(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_LIST)) <> vbNullString Then
 If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "USA" Then
 itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_LIST)) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
 Else
 If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "SPA" Then
 itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_LIST)) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
 Else
 itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = Conversion(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_LIST), MODE_ARGENT, 4)
 End If
3 End If
Else
 itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = Conversion("0", MODE_ARGENT, 4)
End If

itmAchat.ListSubItems(I_COL_ACHAT_PRIX_LIST).Tag = lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PRIX_LIST).Tag

itmAchat.ListSubItems(I_COL_ACHAT_PRIX_LIST).ForeColor = lColor
 
If Trim$(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_NET)) <> vbNullString Then
 If Trim$(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_ESCOMPTE)) <> vbNullString Then
 itmAchat.SubItems(I_COL_ACHAT_ESCOMPTE) = Conversion(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_ESCOMPTE), MODE_POURCENT)
 Else
 itmAchat.SubItems(I_COL_ACHAT_ESCOMPTE) = Conversion("0", MODE_POURCENT)
End If

 If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "USA" Then
 itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_NET)) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
 Else
 If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "SPA" Then
 itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_NET)) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
 Else
 itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_NET), MODE_ARGENT, 4)
 End If
4 End If
4 Else
4 If Trim$(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_SP)) <> vbNullString Then
4 If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "USA" Then
4 itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_SP)) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
4 Else
4 If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "SPA" Then
4 itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_SP)) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
4 Else
4 itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_SP), MODE_ARGENT, 4)
4 End If
4  End If
4  Else
4  itmAchat.SubItems(I_COL_ACHAT_ESCOMPTE) = Conversion("0", MODE_POURCENT)
4  itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion("0", MODE_ARGENT, 4)
4  End If
4  End If

4  itmAchat.ListSubItems(I_COL_ACHAT_ESCOMPTE).ForeColor = lColor
4  itmAchat.ListSubItems(I_COL_ACHAT_PRIX_NET).ForeColor = lColor

50 If lvwfournisseur.SelectedItem.Text = "CHOISIR ULTÉRIEUREMENT" Then
5 itmAchat.SubItems(I_COL_ACHAT_DISTRIB) = " "
 itmAchat.ListSubItems(I_COL_ACHAT_DISTRIB).Tag = 0
 Else
 itmAchat.SubItems(I_COL_ACHAT_DISTRIB) = lvwfournisseur.SelectedItem.Text
 itmAchat.ListSubItems(I_COL_ACHAT_DISTRIB).Tag = lvwfournisseur.SelectedItem.Tag
 End If

 itmAchat.ListSubItems(I_COL_ACHAT_DISTRIB).ForeColor = lColor
 
 'Prix total
 itmAchat.SubItems(I_COL_ACHAT_TOTAL) = Conversion(Round(itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) * itmAchat.Text, 2), MODE_ARGENT)
 itmAchat.ListSubItems(I_COL_ACHAT_TOTAL).ForeColor = lColor
 
 'Calcul des prix
 Call CalculerPrix
 
 'On cache le listview
 frafournisseur.Visible = False

5  Exit Sub

Oups:

5  wOups "frmAchat", "ChoisirFournisseur", Err, Err.number, Err.Description
End Sub

Private Sub ChoisirFournisseurMateriel()

 On Error GoTo Oups

 'On ajoute la pièce en négatif dans le listview
 Dim sQuantite As String
 Dim itmAncien As ListItem
 Dim itmNouveau As ListItem
 
 'Saisie de la quantité
 sQuantite = InputBox("Quelle est la quantité?")

 sQuantite = Replace(sQuantite, ".", ",")
 
 If sQuantite <> vbNullString Then
 If Not IsNumeric(sQuantite) Then
 Call MsgBox("Quantité non numérique!", vbOKOnly, "Erreur")
 
 Exit Sub
 End If
  Else
  Exit Sub
  End If

  If CDbl(sQuantite) <= CDbl(lvwAchat.SelectedItem.Text) Then
  Set itmAncien = lvwAchat.SelectedItem
  Set itmNouveau = lvwAchat.ListItems.Add(itmAncien.Index + 1)
 
  itmNouveau.Checked = itmAncien.Checked
 
  itmNouveau.Text = "-" & sQuantite
    
 'No d'item
itmNouveau.SubItems(I_COL_ACHAT_PIECE) = itmAncien.SubItems(I_COL_ACHAT_PIECE)
 
 'On met la description en francais dans la colonne et la description en anglais
 'dans le tag
1 itmNouveau.SubItems(I_COL_ACHAT_DESCR) = itmAncien.SubItems(I_COL_ACHAT_DESCR)
 itmNouveau.ListSubItems(I_COL_ACHAT_DESCR).Tag = itmAncien.ListSubItems(I_COL_ACHAT_DESCR).Tag
 
 'On met le nom du fabricant dans la colonne et l'ordre de la section dans le tag
 itmNouveau.SubItems(I_COL_ACHAT_MANUFACT) = itmAncien.SubItems(I_COL_ACHAT_MANUFACT)

 'Prix listé
 itmNouveau.SubItems(I_COL_ACHAT_PRIX_LIST) = itmAncien.SubItems(I_COL_ACHAT_PRIX_LIST)

 itmNouveau.ListSubItems(I_COL_ACHAT_PRIX_LIST).Tag = itmAncien.ListSubItems(I_COL_ACHAT_PRIX_LIST).Tag

 itmNouveau.SubItems(I_COL_ACHAT_ESCOMPTE) = itmAncien.SubItems(I_COL_ACHAT_ESCOMPTE)

 itmNouveau.SubItems(I_COL_ACHAT_PRIX_NET) = itmAncien.SubItems(I_COL_ACHAT_PRIX_NET)
 
 'On met le fournisseur dans la colonne et l'id dans le tag
 itmNouveau.SubItems(I_COL_ACHAT_DISTRIB) = lvwfournisseur.SelectedItem.Text
 itmNouveau.ListSubItems(I_COL_ACHAT_DISTRIB).Tag = lvwfournisseur.SelectedItem.Tag
 
 itmNouveau.SubItems(I_COL_ACHAT_TOTAL) = Conversion(Round(CDbl(itmNouveau.Text) * CDbl(itmNouveau.SubItems(I_COL_ACHAT_PRIX_NET)), 2), MODE_ARGENT)

 itmNouveau.SubItems(I_COL_ACHAT_DATE_COMMANDE) = " "
itmNouveau.SubItems(I_COL_ACHAT_DATE_REQUISE) = " "

 itmNouveau.ForeColor = COLOR_BRUN
 itmNouveau.ListSubItems(I_COL_ACHAT_DATE_COMMANDE).ForeColor = COLOR_BRUN
 itmNouveau.ListSubItems(I_COL_ACHAT_DATE_REQUISE).ForeColor = COLOR_BRUN
 itmNouveau.ListSubItems(I_COL_ACHAT_DESCR).ForeColor = COLOR_BRUN
 itmNouveau.ListSubItems(I_COL_ACHAT_DISTRIB).ForeColor = COLOR_BRUN
 itmNouveau.ListSubItems(I_COL_ACHAT_ESCOMPTE).ForeColor = COLOR_BRUN
1  itmNouveau.ListSubItems(I_COL_ACHAT_MANUFACT).ForeColor = COLOR_BRUN
 itmNouveau.ListSubItems(I_COL_ACHAT_PIECE).ForeColor = COLOR_BRUN
 itmNouveau.ListSubItems(I_COL_ACHAT_PRIX_LIST).ForeColor = COLOR_BRUN
 itmNouveau.ListSubItems(I_COL_ACHAT_PRIX_NET).ForeColor = COLOR_BRUN
 itmNouveau.ListSubItems(I_COL_ACHAT_TOTAL).ForeColor = COLOR_BRUN

 'Calcul des prix
 Call CalculerPrix
 
 'On cache le listview
 frafournisseur.Visible = False

 m_bPieceInutile = False
 
 'Resélectionne le premier élément du listview
 If lvwAchat.ListItems.count > 0 Then
 lvwAchat.ListItems(1).Selected = True
 End If
Else
 Call MsgBox("Quantité trop grande!", vbOKOnly, "Erreur")
2  End If

Exit Sub

Oups:

2  wOups "frmAchat", "ChoisirFournisseurMateriel", Err, Err.number, Err.Description
End Sub

Private Sub CalculerPrix()

 On Error GoTo Oups

 Dim dblTotal As Double
 Dim iCompteur As Integer
 
 If lvwAchat.ListItems.count > 0 Then
 For iCompteur = 1 To lvwAchat.ListItems.count
 dblTotal = dblTotal + CDbl(Conversion(lvwAchat.ListItems(iCompteur).SubItems(I_COL_ACHAT_TOTAL), MODE_PAS_FORMAT))
 Next
 
 txtPrixTotal.Text = Conversion(dblTotal, MODE_ARGENT)
 Else
 txtPrixTotal.Text = Conversion(0, MODE_ARGENT)
 End If

  Exit Sub

Oups:

  wOups "frmAchat", "CalculerPrix", Err, Err.number, Err.Description
End Sub

Private Sub lvwFournisseur_LostFocus()

 On Error GoTo Oups

 'On cache le Frame contenant le ListView si le ListView perd le focus
 frafournisseur.Visible = False

 Exit Sub

Oups:

 wOups "frmAchat", "lvwFournisseur_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub lvwPieces_DblClick()

 On Error GoTo Oups

 m_bInventaire = False
 m_bRecherchePiece = -False

 'On affiche lvwFournisseur selon la pièce choisie
 'Rempli les fournisseurs de la pièce choisie
 Call AfficherListeFournisseurs
 
 'Si le listview n'est pas vide
 If lvwfournisseur.ListItems.count = 0 Then
 If MsgBox("Il n'y a aucun fournisseur pour cette pièce!" & vbNewLine & "Voulez-vous en ajouter?", vbYesNo, "Erreur") = vbYes Then
 Screen.MousePointer = vbHourglass
 
 'On ouvre le catalogue sur cet enregistrement
 If m_eCatalogue = ELECTRIQUE Then
 Call FrmCatalogueElec.AfficherForm(cmbCategorie.Text, lvwPieces.SelectedItem.SubItems(I_COL_PIECES_MANUFACT), lvwPieces.SelectedItem.SubItems(I_COL_PIECES_NO_ITEM))
 Else
 Call FrmCatalogueMec.AfficherForm(cmbCategorie.Text, lvwPieces.SelectedItem.SubItems(I_COL_PIECES_MANUFACT), lvwPieces.SelectedItem.SubItems(I_COL_PIECES_NO_ITEM))
  End If
 
  Screen.MousePointer = vbDefault
 
 'On rappelle la méthode
  Call AfficherListeFournisseurs
  End If
  End If

  Exit Sub

Oups:

  wOups "frmAchat", "lvwPieces_DblClick", Err, Err.number, Err.Description
End Sub

Private Sub lvwInventaire_DblClick()

 On Error GoTo Oups

 'On affiche lvwFournisseur selon la pièce choisie
 'Rempli les fournisseurs de la pièce choisie
 
 m_bInventaire = True
 m_bRecherchePiece = False
 
 Call AfficherListeFournisseurs
 
 'Si le listview n'est pas vide
 If lvwfournisseur.ListItems.count = 1 Then
 If MsgBox("Il n'y a aucun fournisseur pour cette pièce!" & vbNewLine & "Voulez-vous en ajouter?", vbYesNo, "Erreur") = vbYes Then
 Screen.MousePointer = vbHourglass
 
 'On ouvre le catalogue sur cet enregistrement
 If m_eCatalogue = ELECTRIQUE Then
 Call FrmCatalogueElec.AfficherForm(cmbCategorie.Text, lvwInventaire.SelectedItem.SubItems(I_COL_INV_MANUFACT), lvwInventaire.SelectedItem.SubItems(I_COL_INV_NO_ITEM))
 Else
 Call FrmCatalogueMec.AfficherForm(cmbCategorie.Text, lvwInventaire.SelectedItem.SubItems(I_COL_INV_MANUFACT), lvwInventaire.SelectedItem.SubItems(I_COL_INV_NO_ITEM))
  End If
 
  Screen.MousePointer = vbDefault
 
 'On rappelle la méthode
  Call AfficherListeFournisseurs
  End If
  End If

  fraInventaire.Visible = False

  Exit Sub

Oups:

  wOups "frmAchat", "lvwInventaire_DblClick", Err, Err.number, Err.Description
End Sub

Private Sub AfficherListeFournisseurs()

 On Error GoTo Oups

 'Méthode qui sert à afficher la liste des fournisseurs
 'Affiche le frame seulement s'il y a des items dans le ListView
 Call RemplirListViewFournisseur
 
 If m_bInventaire = True Then
 m_sQuantite = lvwInventaire.SelectedItem.SubItems(I_COL_INV_QTE_COMMANDE)
 Else
 m_sQuantite = vbNullString
 End If
 
 If lvwfournisseur.ListItems.count > 1 Then
 If m_bRecherchePiece = True Then
 fraPieceTrouve.Visible = False
 End If

  frafournisseur.Visible = True
  Call lvwfournisseur.SetFocus
  End If

  Exit Sub

Oups:

  wOups "frmAchat", "AfficherListeFournisseurs", Err, Err.number, Err.Description
End Sub

Private Sub lvwAchat_KeyDown(KeyCode As Integer, Shift As Integer)

 On Error GoTo Oups

 'S'il est en mode ajout ou modif
 If m_eMode = MODE_AJOUT_MODIF Then
 'Si le listView n'est pas vide
 If lvwAchat.ListItems.count > 0 Then
 'Si la touche pesée est Delete
 If KeyCode = vbKeyDelete Then
 'On l'efface
 Call lvwAchat.ListItems.Remove(lvwAchat.SelectedItem.Index)
 
 Call CalculerPrix
 End If
 End If
 End If

 Exit Sub

Oups:

 wOups "frmAchat", "lvwAchat_KeyDown", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnulerPrix_Click()

 On Error GoTo Oups

 fraPrixPiece.Visible = False

 Exit Sub

Oups:

 wOups "frmAchat", "cmdAnnulerPrix_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdOKPrix_Click()
 'Écrit les prix dans le listview
 On Error GoTo Oups

 Dim rstConfig As ADODB.Recordset
 Dim itmAchat As ListItem
 Dim itmAvant As ListItem
 Dim bPrixSpecial As Boolean
 Dim lColor As Long
 Dim iCompteur As Integer
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

 Set itmAvant = lvwAchat.ListItems(CInt(fraPrixPiece.Tag))
 Set itmAchat = lvwAchat.ListItems.Add(CInt(fraPrixPiece.Tag) + 1)

 lColor = itmAvant.ListSubItems(I_COL_ACHAT_PIECE).ForeColor

itmAchat.Checked = itmAvant.Checked

 'Quantité
 itmAchat.Text = "-" & itmAvant.Text

 'No d'item
itmAchat.SubItems(I_COL_ACHAT_PIECE) = itmAvant.SubItems(I_COL_ACHAT_PIECE)

 'On met la description en francais dans la colonne et la description en anglais
 'dans le tag
 itmAchat.SubItems(I_COL_ACHAT_DESCR) = itmAvant.SubItems(I_COL_ACHAT_DESCR)
itmAchat.ListSubItems(I_COL_ACHAT_DESCR).Tag = itmAvant.ListSubItems(I_COL_ACHAT_DESCR).Tag

 'On met le nom du fabricant dans la col-nne et l'ordre de la section dans le tag
 itmAchat.SubItems(I_COL_ACHAT_MANUFACT) = itmAvant.SubItems(I_COL_ACHAT_MANUFACT)

 'Prix listé
itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = itmAvant.SubItems(I_COL_ACHAT_PRIX_LIST)

 itmAchat.ListSubItems(I_COL_ACHAT_PRIX_LIST).Tag = itmAvant.ListSubItems(I_COL_ACHAT_PRIX_LIST).Tag

itmAchat.SubItems(I_COL_ACHAT_ESCOMPTE) = itmAvant.SubItems(I_COL_ACHAT_ESCOMPTE)

3 itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = itmAvant.SubItems(I_COL_ACHAT_PRIX_NET)

 'On met le fournisseur dans la colonne et l'id dans le tag
 itmAchat.SubItems(I_COL_ACHAT_DISTRIB) = itmAvant.SubItems(I_COL_ACHAT_DISTRIB)
 itmAchat.ListSubItems(I_COL_ACHAT_DISTRIB).Tag = itmAvant.ListSubItems(I_COL_ACHAT_DISTRIB).Tag

 'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
 itmAchat.SubItems(I_COL_ACHAT_TOTAL) = "-" & itmAvant.SubItems(I_COL_ACHAT_TOTAL)

 'Ajout de l'enregistrement avec le nouveau prix
 Set itmAchat = lvwAchat.ListItems.Add(CInt(fraPrixPiece.Tag) + 2)

 itmAchat.Checked = itmAvant.Checked

 'Quantité
 itmAchat.Text = sQuantite

 itmAchat.ForeColor = lColor
 
 'No d'item
 itmAchat.SubItems(I_COL_ACHAT_PIECE) = itmAvant.SubItems(I_COL_ACHAT_PIECE)

 itmAchat.ListSubItems(I_COL_ACHAT_PIECE).ForeColor = lColor

 'On met la description en francais dans la colonne et la description en anglais
 'dans le tag
 itmAchat.SubItems(I_COL_ACHAT_DESCR) = itmAvant.SubItems(I_COL_ACHAT_DESCR)
itmAchat.ListSubItems(I_COL_ACHAT_DESCR).Tag = itmAvant.ListSubItems(I_COL_ACHAT_DESCR).Tag

 itmAchat.ListSubItems(I_COL_ACHAT_DESCR).ForeColor = lColor

 'On met le nom du fabricant dans la colonne et l'ordre de la section dans le tag
itmAchat.SubItems(I_COL_ACHAT_MANUFACT) = itmAvant.SubItems(I_COL_ACHAT_MANUFACT)
 itmAchat.ListSubItems(I_COL_ACHAT_MANUFACT).ForeColor = lColor

If bPrixSpecial = False Then
 If optUSA.Value = True Then
 itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixList.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
 Else
 If optSpain.Value = True Then
4 itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixList.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
4 Else
4 itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = Conversion(txtPrixList.Text, MODE_ARGENT, 4)
4 End If
4 End If

4 itmAchat.ListSubItems(I_COL_ACHAT_PRIX_LIST).Tag = txtPrixList.Text
 
4 itmAchat.ListSubItems(I_COL_ACHAT_PRIX_LIST).ForeColor = lColor
 
 'Escompte
4 If mskEscompte.Text <> vbNullString Then
4 itmAchat.SubItems(I_COL_ACHAT_ESCOMPTE) = Conversion(mskEscompte.Text, MODE_POURCENT)
4 Else
4 itmAchat.SubItems(I_COL_ACHAT_ESCOMPTE) = Conversion("0", MODE_POURCENT)
4  End If

4  itmAchat.ListSubItems(I_COL_ACHAT_ESCOMPTE).ForeColor = lColor

 'Prix net
4  If optUSA.Value = True Then
4  itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixNet.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
4  Else
4  If optSpain.Value = True Then
4  itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixNet.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
4  Else
50 itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(txtPrixNet.Text, MODE_ARGENT, 4)
 End If
 End If

 itmAchat.ListSubItems(I_COL_ACHAT_PRIX_NET).Tag = itmAvant.ListSubItems(I_COL_ACHAT_PRIX_NET).Tag

 itmAchat.ListSubItems(I_COL_ACHAT_PRIX_NET).ForeColor = lColor
 Else
 If optUSA.Value = True Then
 itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
 Else
 If optSpain.Value = True Then
 itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
 Else
5  itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = Conversion(txtPrixSpecial.Text, MODE_ARGENT, 4)
5  End If
5  End If

5  itmAchat.ListSubItems(I_COL_ACHAT_PRIX_LIST).Tag = txtPrixSpecial.Text

5  itmAchat.ListSubItems(I_COL_ACHAT_PRIX_LIST).ForeColor = lColor

5  itmAchat.SubItems(I_COL_ACHAT_ESCOMPTE) = Conversion("0", MODE_POURCENT)

5  itmAchat.ListSubItems(I_COL_ACHAT_ESCOMPTE).ForeColor = lColor

5  If optUSA.Value = True Then
60 itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
  Else
  If optSpain.Value = True Then
  itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
  Else
  itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(txtPrixSpecial.Text, MODE_ARGENT, 4)
  End If
  End If

  itmAchat.ListSubItems(I_COL_ACHAT_PRIX_NET).Tag = itmAvant.ListSubItems(I_COL_ACHAT_PRIX_NET).Tag

  itmAchat.ListSubItems(I_COL_ACHAT_PRIX_NET).ForeColor = lColor
  End If

 'On met le fournisseur dans la colonne et l'id dans le tag
  itmAchat.SubItems(I_COL_ACHAT_DISTRIB) = itmAvant.SubItems(I_COL_ACHAT_DISTRIB)
6  itmAchat.ListSubItems(I_COL_ACHAT_DISTRIB).Tag = itmAvant.ListSubItems(I_COL_ACHAT_DISTRIB).Tag

6  itmAchat.ListSubItems(I_COL_ACHAT_DISTRIB).ForeColor = lColor
 
 'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
6  itmAchat.SubItems(I_COL_ACHAT_TOTAL) = Conversion(CStr(Round(itmAchat.Text * itmAchat.SubItems(I_COL_ACHAT_PRIX_NET), 2)), MODE_ARGENT)

6  itmAchat.ListSubItems(I_COL_ACHAT_TOTAL).ForeColor = lColor

6  itmAchat.SubItems(I_COL_ACHAT_DATE_COMMANDE) = itmAvant.SubItems(I_COL_ACHAT_DATE_COMMANDE)
6  itmAchat.ListSubItems(I_COL_ACHAT_DATE_COMMANDE).ForeColor = lColor

6  itmAchat.SubItems(I_COL_ACHAT_DATE_REQUISE) = itmAvant.SubItems(I_COL_ACHAT_DATE_REQUISE)
6  itmAchat.ListSubItems(I_COL_ACHAT_DATE_REQUISE).ForeColor = lColor

70 If itmAvant.SubItems(I_COL_ACHAT_DATE_COMMANDE) <> "" Then
  itmAvant.ListSubItems(I_COL_ACHAT_DATE_COMMANDE).ForeColor = COLOR_NOIR
  End If

  If itmAvant.SubItems(I_COL_ACHAT_DATE_REQUISE) <> "" Then
  itmAvant.ListSubItems(I_COL_ACHAT_DATE_REQUISE).ForeColor = COLOR_NOIR
  End If

  itmAvant.ListSubItems(I_COL_ACHAT_DESCR).ForeColor = COLOR_NOIR
  itmAvant.ListSubItems(I_COL_ACHAT_DISTRIB).ForeColor = COLOR_NOIR
  itmAvant.ListSubItems(I_COL_ACHAT_ESCOMPTE).ForeColor = COLOR_NOIR
  itmAvant.ListSubItems(I_COL_ACHAT_MANUFACT).ForeColor = COLOR_NOIR
  itmAvant.ListSubItems(I_COL_ACHAT_PIECE).ForeColor = COLOR_NOIR
  itmAvant.ListSubItems(I_COL_ACHAT_PRIX_LIST).ForeColor = COLOR_NOIR
   itmAvant.ListSubItems(I_COL_ACHAT_PRIX_NET).ForeColor = COLOR_NOIR
   itmAvant.ForeColor = COLOR_NOIR
7  itmAvant.ListSubItems(I_COL_ACHAT_TOTAL).ForeColor = COLOR_NOIR

 'Resélectionne le premier élément du listview
7  If lvwAchat.ListItems.count > 0 Then
7  lvwAchat.ListItems(1).Selected = True
7  End If
 
7  m_bMauvaisPrix = False

7  cmbfrs.Locked = False
80 Else
  sPiece = lvwAchat.ListItems(CInt(fraPrixPiece.Tag)).SubItems(I_COL_ACHAT_PIECE)

  For iCompteur = 1 To lvwAchat.ListItems.count
  If lvwAchat.ListItems(iCompteur).SubItems(I_COL_ACHAT_PIECE) = sPiece And lvwAchat.ListItems(iCompteur).ListSubItems(I_COL_ACHAT_PIECE).ForeColor = COLOR_MAGENTA Then
  Set itmAchat = lvwAchat.ListItems(iCompteur)

  itmAchat.ListSubItems(I_COL_ACHAT_PIECE).ForeColor = COLOR_NOIR
  itmAchat.ListSubItems(I_COL_ACHAT_DESCR).ForeColor = COLOR_NOIR
  itmAchat.ListSubItems(I_COL_ACHAT_DISTRIB).ForeColor = COLOR_NOIR
  itmAchat.ListSubItems(I_COL_ACHAT_ESCOMPTE).ForeColor = COLOR_NOIR
  itmAchat.ListSubItems(I_COL_ACHAT_MANUFACT).ForeColor = COLOR_NOIR
  itmAchat.ListSubItems(I_COL_ACHAT_PRIX_LIST).ForeColor = COLOR_NOIR
  itmAchat.ListSubItems(I_COL_ACHAT_PRIX_NET).ForeColor = COLOR_NOIR
   itmAchat.ListSubItems(I_COL_ACHAT_TOTAL).ForeColor = COLOR_NOIR
   itmAchat.ForeColor = COLOR_NOIR

   Call lvwAchat.Refresh
 
   If bPrixSpecial = False Then
 'Prix listé
8  If optUSA.Value = True Then
8  itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixList.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
8  Else
8  If optSpain.Value = True Then
90 itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixList.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
  Else
  itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = Conversion(txtPrixList.Text, MODE_ARGENT, 4)
  End If
  End If

  itmAchat.ListSubItems(I_COL_ACHAT_PRIX_LIST).Tag = txtPrixList.Text
 
 'Escompte
  If mskEscompte.Text <> vbNullString Then
  itmAchat.SubItems(I_COL_ACHAT_ESCOMPTE) = Conversion(mskEscompte.Text, MODE_POURCENT)
  Else
  itmAchat.SubItems(I_COL_ACHAT_ESCOMPTE) = Conversion("0", MODE_POURCENT)
  End If

 'Prix net
  If optUSA.Value = True Then
 itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixNet.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
   Else
 If optSpain.Value = True Then
   itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixNet.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
 Else
   itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(txtPrixNet.Text, MODE_ARGENT, 4)
 End If
9  End If
 Else
 If optUSA.Value = True Then
 itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
 Else
 If optSpain.Value = True Then
 itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
 Else
 itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = Conversion(txtPrixSpecial.Text, MODE_ARGENT, 4)
 End If
 End If

 itmAchat.ListSubItems(I_COL_ACHAT_PRIX_LIST).Tag = txtPrixSpecial.Text
 
 itmAchat.SubItems(I_COL_ACHAT_ESCOMPTE) = Conversion("0", MODE_POURCENT)

10  If optUSA.Value = True Then
10  itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
10  Else
10  If optSpain.Value = True Then
10  itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
10  Else
10  itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(txtPrixSpecial.Text, MODE_ARGENT, 4)
10  End If
1 End If
1 End If

 'On met le fournisseur dans la colonne et l'id dans le tag
1 itmAchat.SubItems(I_COL_ACHAT_DISTRIB) = cmbfrs.LIST(cmbfrs.ListIndex)
 
1 itmAchat.ListSubItems(I_COL_ACHAT_DISTRIB).Tag = cmbfrs.ItemData(cmbfrs.ListIndex)
 
 'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
1 itmAchat.SubItems(I_COL_ACHAT_TOTAL) = Conversion(CStr(Round(Replace(itmAchat.Text, "*", "") * itmAchat.SubItems(I_COL_ACHAT_PRIX_NET), 2)), MODE_ARGENT)

1 If optUSA.Value = True Then
1 itmAchat.ListSubItems(I_COL_ACHAT_TOTAL).Tag = "USA"
1 Else
1 If optSpain.Value = True Then
1 itmAchat.ListSubItems(I_COL_ACHAT_TOTAL).Tag = "SPA"
1 Else
1 itmAchat.ListSubItems(I_COL_ACHAT_TOTAL).Tag = "CAN"
1 End If
1 End If

 End If
1 Next
1 End If

11  Call ModifierPrixCatalogue

1 fraPrixPiece.Visible = False

11  Call CalculerPrix

1 Exit Sub

Oups:

1 wOups "frmAchat", "cmdOKPrix_Click", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboFournisseur()

 On Error GoTo Oups

 Dim rstFRS As ADODB.Recordset

 'Il faut vider le combo avant de le remplir
 Call cmbfrs.Clear

 Set rstFRS = New ADODB.Recordset

 Call rstFRS.Open("SELECT GrbPiecesFRS.*, GrbFournisseur.NomFournisseur FROM GrbPiecesFRS INNER JOIN GrbFournisseur ON GrbPiecesFRS.IDFRS = GrbFournisseur.IDFRS WHERE PIECE = '" & Replace(lvwAchat.SelectedItem.SubItems(I_COL_ACHAT_PIECE), "'", "''") & "' ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)

 'Tant que ce n'est pas la fin des enregistrements
 Do While Not rstFRS.EOF
 Call cmbfrs.AddItem(rstFRS.Fields("NomFournisseur"))

 cmbfrs.ItemData(cmbfrs.newIndex) = rstFRS.Fields("IDFRS")

 Call rstFRS.MoveNext
 Loop

 Exit Sub

Oups:

  wOups "frmAchat", "RemplirComboFournisseur", Err, Err.number, Err.Description
End Sub

Private Sub mnuDateRequise_Click()

 On Error GoTo Oups

 If Trim$(lvwAchat.SelectedItem.SubItems(I_COL_ACHAT_DATE_REQUISE)) = "" Then
 mvwDateRequise.Year = Year(Date)
 mvwDateRequise.Month = Month(Date)
 mvwDateRequise.Day = Day(Date)
 Else
 mvwDateRequise.Year = Left$(lvwAchat.SelectedItem.SubItems(I_COL_ACHAT_DATE_REQUISE), 4)
 mvwDateRequise.Month = Mid$(lvwAchat.SelectedItem.SubItems(I_COL_ACHAT_DATE_REQUISE), 6, 2)
 mvwDateRequise.Day = Right$(lvwAchat.SelectedItem.SubItems(I_COL_ACHAT_DATE_REQUISE), 2)
 End If

 fraDateRequise.Top = lvwAchat.Top

  fraDateRequise.Visible = True

  Exit Sub

Oups:

  wOups "frmAchat", "mnuDateRequise_Click", Err, Err.number, Err.Description
End Sub

Private Sub mvwDateRequise_GotFocus()
 
 On Error GoTo Oups

 m_bMonthViewHasFocus = True

 Exit Sub

Oups:

 wOups "frmAchat", "mvwDateRequise_GotFocus", Err, Err.number, Err.Description
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

  wOups "frmAchat", "txtPrixList_LostFocus", Err, Err.number, Err.Description
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

 wOups "frmAchat", "txtPrixNet_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtPrixNet_GotFocus()

 On Error GoTo Oups

 'Si le prix net prend le focus
 Call CalculerPrixNet

 Exit Sub

Oups:

 wOups "frmAchat", "txtPrixNet_GotFocus", Err, Err.number, Err.Description
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

wOups "frmAchat", "CalculerPrixNet", Err, Err.number, Err.Description
End Sub

Private Sub txtPrixNet_LostFocus()

 On Error GoTo Oups

 txtPrixNet.Text = Replace(txtPrixNet.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmAchat", "txtPrixNet_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub ViderChamps_frs()

 On Error GoTo Oups

 'Vide les champs pieces
 txtPrixList.Text = vbNullString
 mskEscompte.Text = vbNullString
 txtPrixNet.Text = vbNullString
 
 optCAN.Value = True

 Call AfficherDrapeau

 Exit Sub

Oups:

 wOups "frmAchat", "ViderChamps_frs", Err, Err.number, Err.Description
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
 
 'Ouverture du recordset
Call rstPrix.Open("SELECT * FROM GrbPiecesFRS WHERE PIECE = '" & Replace(lvwAchat.ListItems(CInt(fraPrixPiece.Tag)).SubItems(I_COL_ACHAT_PIECE), "'", "''") & "' AND IDFRS = " & cmbfrs.ItemData(cmbfrs.ListIndex), g_connData, adOpenDynamic, adLockOptimistic)
 
rstPrix.Fields("PRIX_LIST") = dblPrixList
rstPrix.Fields("ESCOMPTE") = dblEscompte
rstPrix.Fields("PRIX_NET") = dblPrixNet
rstPrix.Fields("DATE") = ConvertDate(Date)
rstPrix.Fields("ENTRER_PAR") = g_sInitiale
 
If optCAN.Value = True Then
 rstPrix.Fields("DeviseMonétaire") = "CAN"
Else
 If optUSA.Value = True Then
 rstPrix.Fields("DeviseMonétaire") = "USA"
 Else
 rstPrix.Fields("DeviseMonétaire") = "SPA"
 End If
 End If
 
If m_eCatalogue = ELECTRIQUE Then
 rstPrix.Fields("Type") = "E"
1  Else
 rstPrix.Fields("Type") = "M"
 End If
 
Call rstPrix.Update
 
Call rstPrix.Close
Set rstPrix = Nothing
 
Exit Sub

Oups:

wOups "frmAchat", "ModifierPrixCatalogue", Err, Err.number, Err.Description
End Sub

Private Sub optCAN_Click()

 On Error GoTo Oups

 'Dépendant la devise, affiche le drapeau
 Call AfficherDrapeau

 Exit Sub

Oups:

 wOups "frmAchat", "optCAN_Click", Err, Err.number, Err.Description
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

  wOups "frmAchat", "AfficherDrapeau", Err, Err.number, Err.Description
End Sub

Private Sub optSpain_Click()

 On Error GoTo Oups

 'dependant la devise, affiche le drapeau
 Call AfficherDrapeau

 Exit Sub

Oups:

 wOups "frmAchat", "optSpain_Click", Err, Err.number, Err.Description
End Sub

Private Sub optUSA_Click()

 On Error GoTo Oups

 'dependant la devise, affiche le drapeau
 Call AfficherDrapeau

 Exit Sub

Oups:

 wOups "frmAchat", "optUSA_Click", Err, Err.number, Err.Description
End Sub

Private Sub mskEscompte_GotFocus()

 On Error GoTo Oups

 'Quand le maskEdit prend le focus, on set le masque
 If mskEscompte.Enabled = True Then
 mskEscompte.mask = "0,####"
 End If

 Exit Sub

Oups:

 wOups "frmAchat", "mskEscompte_GotFocus", Err, Err.number, Err.Description
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

 wOups "frmAchat", "mskEscompte_LostFocus", Err, Err.number, Err.Description
End Sub

Public Sub Commande()

 On Error GoTo Oups

 Dim rstPiece As ADODB.Recordset
 Dim rstBCPiece As ADODB.Recordset
 Dim rstBC As ADODB.Recordset
 Dim rstFRS As ADODB.Recordset
 Dim iIDFRS As Integer
 Dim sFRS As String
 Dim sNoBC As String
 Dim sWhere As String
 Dim sWherePiece As String
 Dim sWhereNoLigne As String
  Dim sDateRequise As String
  Dim sNoLigne As String
  Dim bPremier As Boolean
  Dim bPremierNoLigne As Boolean

  sFRS = DR_Commande.Sections("Section2").Controls("lblFournisseur").Caption
  sNoBC = DR_Commande.Sections("Section2").Controls("lblNoBC").Caption

  Set rstBC = New ADODB.Recordset
  Set rstFRS = New ADODB.Recordset
10 Set rstPiece = New ADODB.Recordset
Set rstBCPiece = New ADODB.Recordset

Call rstBC.Open("SELECT * FROM GrbBonsCommandes WHERE NoBonCommande = '" & sNoBC & "'", g_connData, adOpenDynamic, adLockOptimistic)

Do While Not rstBC.EOF
 Call rstFRS.Open("SELECT IDFRS, NomFournisseur FROM GrbFournisseur WHERE IDFRS = " & rstBC.Fields("NoFournisseur"), g_connData, adOpenDynamic, adLockOptimistic)

 If rstFRS.Fields("NomFournisseur") = sFRS Then
 iIDFRS = rstFRS.Fields("IDFRS")

 sDateRequise = rstBC.Fields("DateRequise")

 Call rstFRS.Close

 Exit Do
 End If

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
 sWhere = "(IDAchat = '" & Left$(txtNoAchat.Text, 9) & "' AND IndexAchat = " & Int(Right$(txtNoAchat.Text, 3)) & ")"

1  sWherePiece = "PIECE In ("
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
 Else
 Do While InStr(1, sNoLigne, ",") > 0
 sWherePiece = sWherePiece & ", '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
 sWhereNoLigne = sWhereNoLigne & ", " & Left$(sNoLigne, InStr(1, sNoLigne, ",") - 1)

 sNoLigne = Right$(sNoLigne, Len(sNoLigne) - (InStr(1, sNoLigne, ",") + 1))
 Loop

 If Trim$(sNoLigne) <> "" Then
 sWherePiece = sWherePiece & ", '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
 sWhereNoLigne = sWhereNoLigne & ", " & sNoLigne
4 End If
4 End If
4 End If
4 End If
 
4 Call rstBCPiece.MoveNext
4 Loop

4 sWherePiece = sWherePiece & ")"
4 sWhereNoLigne = sWhereNoLigne & ")"

4 sWhere = sWhere & " AND " & sWherePiece & " AND " & sWhereNoLigne
 
4 Call rstBCPiece.Close
4 Set rstBCPiece = Nothing

4  Call rstPiece.Open("SELECT * FROM GrbAchat_Pieces WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
 
4  Do While Not rstPiece.EOF
4  rstPiece.Fields("Commandé") = True

4  rstPiece.Fields("DateCommande") = ConvertDate(Date)

4  rstPiece.Fields("DateRequise") = sDateRequise
 
4  Call rstPiece.Update
 
4  Call rstPiece.MoveNext
4  Loop
 
50 Call rstPiece.Close
50 Set rstPiece = Nothing
 
 Call RemplirListViewAchat

 Exit Sub

Oups:

 wOups "frmAchat", "Commande", Err, Err.number, Err.Description
End Sub

Private Sub cmdRetour_Click()

 On Error GoTo Oups

 Dim sIDAchat As String
 Dim iIndexAchat As Integer
 Dim rstAchat As ADODB.Recordset

 If cmbNoAchat.ListCount > 0 Then
 sIDAchat = Trim$(Left$(txtNoAchat.Text, InStr(1, txtNoAchat.Text, "-") + 2))
 iIndexAchat = CInt(Trim$(Right$(txtNoAchat.Text, Len(txtNoAchat.Text) - InStrRev(txtNoAchat.Text, "-"))))

 Set rstAchat = New ADODB.Recordset

 Call rstAchat.Open("SELECT Modification, Par FROM GrbAchat WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat, g_connData, adOpenDynamic, adLockOptimistic)

 If rstAchat.Fields("Modification") = False Then
 Call rstAchat.Close
  Set rstAchat = Nothing

  Screen.MousePointer = vbHourglass
 
  iIndexAchat = CInt(Right$(txtNoAchat.Text, 3))

  Call frmRetourMarchandise.AfficherAchat(sIDAchat, iIndexAchat, g_sUserID)

  Call cmbNoAchat_Click

  Screen.MousePointer = vbDefault
  Else
  Call MsgBox("Cet achat est en modification par " & rstAchat.Fields("Par") & "!", vbOKOnly, "Erreur")

 Call rstAchat.Close
Set rstAchat = Nothing
 End If
End If

Exit Sub

Oups:

wOups "frmAchat", "cmdRetour_Click", Err, Err.number, Err.Description
End Sub

Private Sub OuvrirAchat(ByVal bOuvrir As Boolean)
 'Remplit ou vide les champs Modification et Par
 On Error GoTo Oups

 Dim rstAchat As ADODB.Recordset
 Dim sIDAchat As String
 Dim iIndexAchat As Integer

 sIDAchat = Trim$(Left$(txtNoAchat.Text, InStr(1, txtNoAchat.Text, "-") + 2))
 iIndexAchat = CInt(Trim$(Right$(txtNoAchat.Text, Len(txtNoAchat.Text) - InStrRev(txtNoAchat.Text, "-"))))

 Set rstAchat = New ADODB.Recordset

 rstAchat.CursorLocation = adUseServer

 Call rstAchat.Open("SELECT Modification, Par FROM GrbAchat WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat, g_connData, adOpenDynamic, adLockOptimistic)

 Do While Not rstAchat.EOF
 If bOuvrir = True Then
  rstAchat.Fields("Modification") = True
  rstAchat.Fields("Par") = g_sEmploye
  Else
  rstAchat.Fields("Modification") = False
  rstAchat.Fields("Par") = ""
  End If

  Call rstAchat.Update

  Call rstAchat.MoveNext
10 Loop

Call rstAchat.Close
Set rstAchat = Nothing

Exit Sub

Oups:

wOups "frmAchat", "OuvrirAchat", Err, Err.number, Err.Description
End Sub

Private Sub lvwPieceTrouve_DblClick()

 On Error GoTo Oups

 Dim iCompteur As Integer

 m_bRecherchePiece = True
 m_bInventaire = False

 'On affiche lvwFournisseur selon la pièce choisie
 'Rempli les fournisseurs de la pièce choisie
 Call AfficherListeFournisseurs
 
 'si le listview n'est pas vide
 If lvwfournisseur.ListItems.count = 1 Then
 If MsgBox("Il n'y a aucun fournisseur pour cette pièce!" & vbNewLine & "Voulez-vous en ajouter?", vbYesNo, "Erreur") = vbYes Then
 Screen.MousePointer = vbHourglass
 
 'On ouvre le catalogue sur cet enregistrement
 If m_eCatalogue = ELECTRIQUE Then
 Call FrmCatalogueElec.AfficherForm(lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_CATEGORIE), lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_MANUFACT), lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM))
 Else
  Call FrmCatalogueMec.AfficherForm(lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_CATEGORIE), lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_MANUFACT), lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM))
  End If

  Screen.MousePointer = vbDefault
 
 'On rappelle la méthode
  Call AfficherListeFournisseurs
  End If
  End If

  Exit Sub

Oups:

  wOups "frmAchat", "lvwPieceTrouve_DblClick", Err, Err.number, Err.Description
End Sub

Private Sub cmdOKPieceTrouve_Click()

 On Error GoTo Oups

 m_bRecherchePiece = True
 m_bInventaire = False

 'On affiche lvwFournisseur selon la pièce choisie
 'Rempli les fournisseurs de la pièce choisie
 Call AfficherListeFournisseurs
 
 'si le listview n'est pas vide
 If lvwfournisseur.ListItems.count = 1 Then
 If MsgBox("Il n'y a aucun fournisseur pour cette pièce!" & vbNewLine & "Voulez-vous en ajouter?", vbYesNo, "Erreur") = vbYes Then
 Screen.MousePointer = vbHourglass
 
 'On ouvre le catalogue sur cet enregistrement
 If m_eCatalogue = ELECTRIQUE Then
 Call FrmCatalogueElec.AfficherForm(lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_CATEGORIE), lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_MANUFACT), lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM))
 Else
 Call FrmCatalogueMec.AfficherForm(lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_CATEGORIE), lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_MANUFACT), lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM))
  End If
 
  Screen.MousePointer = vbDefault
 
 'On rappelle la méthode
  Call AfficherListeFournisseurs
  End If
  End If

  Exit Sub

Oups:

  wOups "frmAchat", "cmdOKPieceTrouve_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnulerPieceTrouve_Click()

 On Error GoTo Oups

 fraPieceTrouve.Visible = False

 Exit Sub

Oups:

 wOups "frmAchat", "cmdAnnulerPieceTrouve", Err, Err.number, Err.Description
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

If m_eCatalogue = ELECTRIQUE Then
 If iIndexColumn = I_COL_PIECES_NO_ITEM Then
 Call rstPiece.Open("SELECT * FROM GrbCatalogueElec WHERE INSTR(1, " & sChamps & ", '" & sRecherche & "') > 0", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstPiece.Open("SELECT * FROM GrbCatalogueElec WHERE INSTR(1, " & sChamps & ", '" & sTexte & "') > 0", g_connData, adOpenDynamic, adLockOptimistic)
 End If
Else
 If iIndexColumn = I_COL_PIECES_NO_ITEM Then
 Call rstPiece.Open("SELECT * FROM GrbCatalogueMec WHERE INSTR(1, " & sChamps & ", '" & sRecherche & "') > 0", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstPiece.Open("SELECT * FROM GrbCatalogueMec WHERE INSTR(1, " & sChamps & ", '" & sTexte & "') > 0", g_connData, adOpenDynamic, adLockOptimistic)
 End If
 End If

 'Pour chaque enregistrement
Do While Not rstPiece.EOF
 'On ajoute dans le ListView
 Set itmPiece = lvwPieceTrouve.ListItems.Add

1  If m_eCatalogue = ELECTRIQUE Then
 If Not IsNull(rstPiece.Fields("TEMPS")) Then
 itmPiece.Tag = rstPiece.Fields("TEMPS")
 Else
 itmPiece.Tag = vbNullString
 End If
 End If

 If Not IsNull(rstPiece.Fields("PIECE_GRB")) Then
 itmPiece.Text = rstPiece.Fields("PIECE_GRB")
 Else
 itmPiece.Text = ""
 End If

 itmPiece.SubItems(I_COL_RECH_NO_ITEM) = rstPiece.Fields("PIECE")
itmPiece.SubItems(I_COL_RECH_CATEGORIE) = cmbCategorie.LIST(iCompteur)

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

Call rstPiece.Close
Set rstPiece = Nothing

3  Exit Sub

Oups:

wOups "frmAchat", "RemplirListViewRecherche", Err, Err.number, Err.Description
End Sub

Private Sub cmdMaterielInutile_Click()

 On Error GoTo Oups

 Dim itmAchat As ListItem

 If lvwAchat.ListItems.count > 0 Then
 Set itmAchat = lvwAchat.SelectedItem
 
 'Si la quantité est plus grande que 0
 If CDbl(itmAchat.Text) > 0 Then
 m_bPieceInutile = True
 m_bRecherchePiece = False

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

  Exit Sub

Oups:

10 wOups "frmAchat", "cmdMaterielInutile_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdMauvaisPrix_Click()

 On Error GoTo Oups

 Dim iCompteur As Integer
 Dim itmAchat As ListItem

 If lvwAchat.ListItems.count > 0 Then
 Set itmAchat = lvwAchat.SelectedItem
 
 'Si la quantité est plus grande que 0
 If CDbl(itmAchat.Text) > 0 Then
 Call ViderChamps_frs

 Call RemplirComboFournisseur

 For iCompteur = 0 To cmbfrs.ListCount - 1
 If cmbfrs.ItemData(iCompteur) = itmAchat.ListSubItems(I_COL_ACHAT_DISTRIB).Tag Then
 cmbfrs.ListIndex = iCompteur

  Exit For
  End If
  Next

  cmbfrs.Locked = True

  fraPrixPiece.Tag = itmAchat.Index

  m_bMauvaisPrix = True

  fraPrixPiece.Visible = True
  Else
 Call MsgBox("La quantité est déjà dans le négatif!", vbOKOnly, "Erreur")
1 End If
End If

Exit Sub

Oups:

wOups "frmAchat", "cmdMauvaisPrix_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdReception_Click()

 On Error GoTo Oups

 Dim iCompteur As Integer
 Dim bOuvert As Boolean

 If m_eCatalogue = ELECTRIQUE Then
 For iCompteur = 0 To Forms.count - 1
 If Forms(iCompteur).Name = "FrmReceptionElec" Then
 bOuvert = True

 Exit For
 End If
 Next

 If bOuvert = True Then
  Call Unload(FrmReceptionElec)
  End If

  Call FrmReceptionElec.AfficherAchat(g_sUserID, txtNoAchat.Text)

  Call RemplirListViewAchat
  Else
  For iCompteur = 0 To Forms.count - 1
  If Forms(iCompteur).Name = "FrmReceptionMec" Then
  bOuvert = True

 Exit For
End If
 Next

 If bOuvert = True Then
 Call Unload(FrmReceptionMec)
 End If

 Call FrmReceptionMec.AfficherAchat(g_sUserID, txtNoAchat.Text)

 Call RemplirListViewAchat
End If

Exit Sub

Oups:

wOups "frmAchat", "cmdReception_Click", Err, Err.number, Err.Description
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

  wOups "frmAchat", "txtPrixSpecial_Change", Err, Err.number, Err.Description
End Sub

Private Sub txtPrixSpecial_LostFocus()

 On Error GoTo Oups

 txtPrixSpecial.Text = Replace(txtPrixSpecial.Text, ".", ",")

 Exit Sub

Oups:

 wOups "frmAchat", "txtPrixSpecial_LostFocus", Err, Err.number, Err.Description
End Sub
