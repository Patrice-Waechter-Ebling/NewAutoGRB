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
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   0
         Appearance      =   1
         StartOfWeek     =   90243073
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
Private Const I_COL_ACHAT_QUANTITE      As Integer = 0
Private Const I_COL_ACHAT_PIECE         As Integer = 1
Private Const I_COL_ACHAT_DESCR         As Integer = 2
Private Const I_COL_ACHAT_MANUFACT      As Integer = 3
Private Const I_COL_ACHAT_PRIX_LIST     As Integer = 4
Private Const I_COL_ACHAT_ESCOMPTE      As Integer = 5
Private Const I_COL_ACHAT_PRIX_NET      As Integer = 6
Private Const I_COL_ACHAT_DISTRIB       As Integer = 7
Private Const I_COL_ACHAT_TOTAL         As Integer = 8
Private Const I_COL_ACHAT_DATE_COMMANDE As Integer = 9
Private Const I_COL_ACHAT_DATE_REQUISE  As Integer = 10

'Index des colonnes de lvwPieces
Private Const I_COL_PIECES_PIECE_GRB   As Integer = 0
Private Const I_COL_PIECES_NO_ITEM     As Integer = 1
Private Const I_COL_PIECES_MANUFACT    As Integer = 2
Private Const I_COL_PIECES_DESCR_FR    As Integer = 3
Private Const I_COL_PIECES_DESCR_EN    As Integer = 4
Private Const I_COL_PIECES_COMMENT     As Integer = 5

'Index des colonnes de lvwPieceTrouve
Private Const I_COL_RECH_PIECE_GRB     As Integer = 0
Private Const I_COL_RECH_NO_ITEM       As Integer = 1
Private Const I_COL_RECH_CATEGORIE     As Integer = 2
Private Const I_COL_RECH_MANUFACT      As Integer = 3
Private Const I_COL_RECH_DESCR_FR      As Integer = 4
Private Const I_COL_RECH_DESCR_EN      As Integer = 5

'Index des colonnes de lvwInventaire
Private Const I_COL_INV_NO_ITEM        As Integer = 0
Private Const I_COL_INV_MANUFACT       As Integer = 1
Private Const I_COL_INV_DESCR          As Integer = 2
Private Const I_COL_INV_COMMENT        As Integer = 3
Private Const I_COL_INV_QTE_STOCK      As Integer = 4
Private Const I_COL_INV_QTE_MINIMUM    As Integer = 5
Private Const I_COL_INV_QTE_COMMANDE   As Integer = 6

'Index des colonnes de lvwFournisseur
Private Const I_COL_FRS_FRS            As Integer = 0
Private Const I_COL_FRS_PERS_RESS      As Integer = 1
Private Const I_COL_FRS_DATE           As Integer = 2
Private Const I_COL_FRS_ENTRER_PAR     As Integer = 3
Private Const I_COL_FRS_VALIDE         As Integer = 4
Private Const I_COL_FRS_PRIX_LIST      As Integer = 5
Private Const I_COL_FRS_ESCOMPTE       As Integer = 6
Private Const I_COL_FRS_PRIX_NET       As Integer = 7
Private Const I_COL_FRS_PRIX_SP        As Integer = 8
Private Const I_COL_FRS_QUOTER         As Integer = 9

'Index de cmbTri
Private Const I_CMB_PIECE_GRB          As Integer = 0
Private Const I_CMB_PIECE              As Integer = 1
Private Const I_CMB_FABRICANT          As Integer = 2
Private Const I_CMB_DESCR_FR           As Integer = 3
Private Const I_CMB_DESCR_EN           As Integer = 4

'Valeur servant au grandeur du lvwSoumission si en mode modif ou inactif
Private Const I_TOP_AJOUT_MODIF        As Integer = 4320
Private Const I_HEIGHT_AJOUT_MODIF     As Integer = 2535
Private Const I_TOP_INACTIF            As Integer = 1680
Private Const I_HEIGHT_INACTIF         As Integer = 5175

'Width de cmbCategorie
Private Const I_WIDTH_CATEGORIE_ELEC   As Integer = 1695
Private Const I_WIDTH_CATEGORIE_MEC    As Integer = 5175

'Énumération servant à savoir si le form est en mode modif/ajout ou en mode
'inactif (affichage seulement)
Private Enum enumMode
  MODE_AJOUT_MODIF = 0
  MODE_INACTIF = 1
End Enum

'Pour la recherche de pièce dans lvwPieces
Private m_sTri               As String

'Pour savoir quelle colonne trier
Private m_iCol               As Integer

'Modes du form
Private m_bModeAjout         As Boolean
Private m_bModeAffichage     As Boolean

'Pour pouvoir afficher le dernier achat visualisé
Private m_sAncienAchat       As String

Public m_sNoAchat            As String

Public m_bAnnuler            As Boolean

Private m_eMode              As enumMode

Private m_eCatalogue         As enumCatalogue

Private m_bInventaire        As Boolean

Private m_sQuantite          As String

Private m_bRecherchePiece    As Boolean

Private m_bPieceInutile      As Boolean

'Pour savoir si le changement de prix a été appelé à partir
'du bouton "Mauvais Prix"
Private m_bMauvaisPrix       As Boolean

Private m_bMonthViewHasFocus As Boolean

Private Sub AnnulerCommande()

5       On Error GoTo AfficherErreur

10      Dim itmAvant      As ListItem
15      Dim itmAnnulation As ListItem

20      Set itmAvant = lvwAchat.SelectedItem
25      Set itmAnnulation = lvwAchat.ListItems.Add(itmAvant.Index + 1)

30      itmAnnulation.Checked = itmAvant.Checked

        'Quantité
35      itmAnnulation.Text = "-" & itmAvant.Text

        'On met l'id de la section dans le tag du listItem
40      itmAnnulation.Tag = itmAvant.Tag

        'No d'item
45      itmAnnulation.SubItems(I_COL_ACHAT_PIECE) = itmAvant.SubItems(I_COL_ACHAT_PIECE)

        'On met le nom de la sous-section dans le tag du no d'item
50      itmAnnulation.ListSubItems(I_COL_ACHAT_PIECE).Tag = itmAvant.ListSubItems(I_COL_ACHAT_PIECE).Tag

        'On met la description en francais dans la colonne et la description en anglais
        'dans le tag
55      itmAnnulation.SubItems(I_COL_ACHAT_DESCR) = itmAvant.SubItems(I_COL_ACHAT_DESCR)
60      itmAnnulation.ListSubItems(I_COL_ACHAT_DESCR).Tag = itmAvant.ListSubItems(I_COL_ACHAT_DESCR).Tag

        'On met le nom du fabricant dans la col-nne et l'ordre de la section dans le tag
65      itmAnnulation.SubItems(I_COL_ACHAT_MANUFACT) = itmAvant.SubItems(I_COL_ACHAT_MANUFACT)
70      itmAnnulation.ListSubItems(I_COL_ACHAT_MANUFACT).Tag = itmAvant.ListSubItems(I_COL_ACHAT_MANUFACT).Tag

        'Prix listé
75      itmAnnulation.SubItems(I_COL_ACHAT_PRIX_LIST) = itmAvant.SubItems(I_COL_ACHAT_PRIX_LIST)

80      itmAnnulation.ListSubItems(I_COL_ACHAT_PRIX_LIST).Tag = itmAvant.ListSubItems(I_COL_ACHAT_PRIX_LIST).Tag

85      itmAnnulation.SubItems(I_COL_ACHAT_ESCOMPTE) = itmAvant.SubItems(I_COL_ACHAT_ESCOMPTE)

90      itmAnnulation.SubItems(I_COL_ACHAT_PRIX_NET) = itmAvant.SubItems(I_COL_ACHAT_PRIX_NET)

        'On met le fournisseur dans la colonne et l'id dans le tag
95      itmAnnulation.SubItems(I_COL_ACHAT_DISTRIB) = itmAvant.SubItems(I_COL_ACHAT_DISTRIB)
100     itmAnnulation.ListSubItems(I_COL_ACHAT_DISTRIB).Tag = itmAvant.ListSubItems(I_COL_ACHAT_DISTRIB).Tag

        'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
105     itmAnnulation.SubItems(I_COL_ACHAT_TOTAL) = "-" & itmAvant.SubItems(I_COL_ACHAT_TOTAL)

110     itmAnnulation.ForeColor = COLOR_VERT_FORET
115     itmAnnulation.ListSubItems(I_COL_ACHAT_PIECE).ForeColor = COLOR_VERT_FORET
120     itmAnnulation.ListSubItems(I_COL_ACHAT_DESCR).ForeColor = COLOR_VERT_FORET
125     itmAnnulation.ListSubItems(I_COL_ACHAT_DISTRIB).ForeColor = COLOR_VERT_FORET
130     itmAnnulation.ListSubItems(I_COL_ACHAT_ESCOMPTE).ForeColor = COLOR_VERT_FORET
135     itmAnnulation.ListSubItems(I_COL_ACHAT_MANUFACT).ForeColor = COLOR_VERT_FORET
140     itmAnnulation.ListSubItems(I_COL_ACHAT_PRIX_LIST).ForeColor = COLOR_VERT_FORET
145     itmAnnulation.ListSubItems(I_COL_ACHAT_PRIX_NET).ForeColor = COLOR_VERT_FORET
150     itmAnnulation.ListSubItems(I_COL_ACHAT_TOTAL).ForeColor = COLOR_VERT_FORET
                      
155     itmAnnulation.Bold = True
160     itmAnnulation.ListSubItems(I_COL_ACHAT_PIECE).Bold = True
165     itmAnnulation.ListSubItems(I_COL_ACHAT_DESCR).Bold = True
170     itmAnnulation.ListSubItems(I_COL_ACHAT_DISTRIB).Bold = True
175     itmAnnulation.ListSubItems(I_COL_ACHAT_ESCOMPTE).Bold = True
180     itmAnnulation.ListSubItems(I_COL_ACHAT_MANUFACT).Bold = True
185     itmAnnulation.ListSubItems(I_COL_ACHAT_PRIX_LIST).Bold = True
190     itmAnnulation.ListSubItems(I_COL_ACHAT_PRIX_NET).Bold = True
195     itmAnnulation.ListSubItems(I_COL_ACHAT_TOTAL).Bold = True

200     itmAvant.ForeColor = COLOR_NOIR
205     itmAvant.ListSubItems(I_COL_ACHAT_PIECE).ForeColor = COLOR_NOIR
210     itmAvant.ListSubItems(I_COL_ACHAT_DESCR).ForeColor = COLOR_NOIR
215     itmAvant.ListSubItems(I_COL_ACHAT_DISTRIB).ForeColor = COLOR_NOIR
220     itmAvant.ListSubItems(I_COL_ACHAT_ESCOMPTE).ForeColor = COLOR_NOIR
225     itmAvant.ListSubItems(I_COL_ACHAT_MANUFACT).ForeColor = COLOR_NOIR
230     itmAvant.ListSubItems(I_COL_ACHAT_PRIX_LIST).ForeColor = COLOR_NOIR
235     itmAvant.ListSubItems(I_COL_ACHAT_PRIX_NET).ForeColor = COLOR_NOIR
240     itmAvant.ListSubItems(I_COL_ACHAT_TOTAL).ForeColor = COLOR_NOIR
245     itmAvant.ListSubItems(I_COL_ACHAT_DATE_COMMANDE).ForeColor = COLOR_NOIR
250     itmAvant.ListSubItems(I_COL_ACHAT_DATE_REQUISE).ForeColor = COLOR_NOIR

255     Call lvwAchat.Refresh

260     Call CalculerPrix

265     Exit Sub

AfficherErreur:

270     woups "frmAchat", "AnnulerCommande", Err, Erl
End Sub

Public Sub Afficher(ByVal eCatalogue As enumCatalogue)

5       On Error GoTo AfficherErreur

10      Call Unload(frmChoixProjSoum)
  
15      m_eCatalogue = eCatalogue
  
20      Select Case eCatalogue
          'Si c'est électrique
          Case ELECTRIQUE:
25          Me.Caption = "Achat électrique"
30          cmbCategorie.width = I_WIDTH_CATEGORIE_ELEC
      
          Case MECANIQUE:
35          Me.Caption = "Achat mécanique"
40          cmbCategorie.width = I_WIDTH_CATEGORIE_MEC
45      End Select
  
        'Initialise le tri à PIECE_GRB
50      cmbTri.ListIndex = I_CMB_PIECE
  
55      Call RemplirComboAchat(vbNullString)
  
        'Rempli le combo des catégories de pièce
60      Call RemplirComboCategorie
    
65      Call AfficherControles(MODE_INACTIF)
  
70      Call Me.Show

75      Exit Sub

AfficherErreur:

80      woups "frmAchat", "Afficher", Err, Erl
End Sub

Private Sub AfficherControles(ByVal eMode As enumMode)

5       On Error GoTo AfficherErreur
        
        'Affichage des boutons selon si c'est un ajout/modif ou un affichage
10      Dim bAjouter      As Boolean
15      Dim bModifier     As Boolean
20      Dim bSupprimer    As Boolean
25      Dim bEnregistrer  As Boolean
30      Dim bAnnuler      As Boolean
35      Dim bFermer       As Boolean
40      Dim bImprimer     As Boolean
45      Dim bBonCommande  As Boolean
50      Dim bTri          As Boolean
55      Dim bCmbAchat     As Boolean
60      Dim bDemandePrix  As Boolean
65      Dim bRetour       As Boolean
70      Dim bInventaire   As Boolean
75      Dim bInutile      As Boolean
80      Dim bPrix         As Boolean
85      Dim bPieces       As Boolean
90      Dim bReception    As Boolean
    
95      m_eMode = eMode
  
100     Select Case eMode
          Case MODE_AJOUT_MODIF:
105         bEnregistrer = True
110         bAnnuler = True
115         bPieces = True
120         bTri = True
125         bInventaire = True
130         bPrix = True
135         bInutile = True
             
          Case MODE_INACTIF:
140         bModifier = True
145         bFermer = True
150         bImprimer = True
155         bAjouter = True
160         bBonCommande = True
165         bSupprimer = True
170         bCmbAchat = True
175         bDemandePrix = True
180         bRetour = True
185         bReception = True
190     End Select
  
195     cmbNoAchat.Visible = bCmbAchat
200     txtNoAchat.Visible = Not bCmbAchat
  
205     Cmdajouter.Visible = bAjouter
210     cmdModifier.Visible = bModifier
215     cmdsupprimer.Visible = bSupprimer
220     cmdEnregistrer.Visible = bEnregistrer
225     cmdAnnuler.Visible = bAnnuler
230     Cmdfermer.Visible = bFermer
235     cmdImprimer.Visible = bImprimer
240     cmdBonCommande.Visible = bBonCommande
245     cmdDemande.Visible = bDemandePrix
250     cmdRetour.Visible = bRetour
255     cmdInventaire.Visible = bInventaire
260     cmdReception.Visible = bReception
265     lblCategorie.Visible = bPieces
270     cmbCategorie.Visible = bPieces
275     lvwPieces.Visible = bPieces

280     lblTri.Visible = bTri
285     cmbTri.Visible = bTri
290     cmdTri.Visible = bTri
295     cmdRafraichir.Visible = bTri
         
        'Exception puisqu'il y en a qu'un seul
300     If m_eMode = MODE_AJOUT_MODIF Then
305       txtRaison.Locked = False
310     Else
315       txtRaison.Locked = True
320     End If

325     If m_eMode = MODE_AJOUT_MODIF Then
330       lvwAchat.Top = I_TOP_AJOUT_MODIF
335       lvwAchat.Height = I_HEIGHT_AJOUT_MODIF
340     Else
345       lvwAchat.Top = I_TOP_INACTIF
350       lvwAchat.Height = I_HEIGHT_INACTIF
355     End If

360     Exit Sub

AfficherErreur:

365     woups "frmAchat", "AfficherControles", Err, Erl
End Sub

Private Sub cmbCategorie_Click()

5       On Error GoTo AfficherErreur

        'Rempli lvwPieces selon la catégorie de pièce choisie
10      Call RemplirListViewPieces

15      Exit Sub

AfficherErreur:

20      woups "frmAchat", "cmbCategorie_Click", Err, Erl
End Sub

Private Sub cmbNoAchat_Click()

5       On Error GoTo AfficherErreur

10      Dim sNomClient  As String
15      Dim sNomContact As String
20      Dim iCompteur   As Integer
25      Dim rstAchat    As ADODB.Recordset
30      Dim sIDAchat    As String
35      Dim iIndexAchat As Integer
  
40      Screen.MousePointer = vbHourglass
  
45      txtNoAchat.Text = cmbNoAchat.Text
  
        'Rempli les valeurs de l'achat sélectionné
50      Call RemplirAchat

55      sIDAchat = Trim$(Left$(txtNoAchat.Text, InStr(1, txtNoAchat.Text, "-") + 2))
60      iIndexAchat = CInt(Trim$(Right$(txtNoAchat.Text, Len(txtNoAchat.Text) - InStrRev(txtNoAchat.Text, "-"))))

65      Set rstAchat = New ADODB.Recordset

70      Call rstAchat.Open("SELECT Modification, Par FROM GRB_Achat WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat, g_connData, adOpenDynamic, adLockOptimistic)

75      If rstAchat.Fields("Modification") = True And rstAchat.Fields("Par") = g_sEmploye Then
80        cmdReset.Visible = True
85      Else
90        cmdReset.Visible = False
95      End If

100     Call rstAchat.Close
105     Set rstAchat = Nothing
       
110     Screen.MousePointer = vbDefault

115     Exit Sub

AfficherErreur:

120     woups "frmAchat", "cmbNoAchat_Click", Err, Erl
End Sub

Private Sub RemplirAchat()

5       On Error GoTo AfficherErreur

10      Dim rstAchat    As ADODB.Recordset
15      Dim rstEmploye  As ADODB.Recordset
20      Dim sNoAchat    As String
25      Dim iIndexAchat As Integer
  
30      sNoAchat = Left$(txtNoAchat.Text, 9)
  
35      iIndexAchat = CInt(Right$(txtNoAchat.Text, 3))
  
40      Set rstAchat = New ADODB.Recordset
45      Set rstEmploye = New ADODB.Recordset
  
50      Call rstAchat.Open("SELECT * FROM GRB_Achat WHERE IDAchat = '" & sNoAchat & "' AND IndexAchat = " & iIndexAchat, g_connData, adOpenDynamic, adLockOptimistic)
  
55      Call rstEmploye.Open("SELECT Employe FROM GRB_Employés WHERE noEmploye = " & rstAchat.Fields("Acheteur"), g_connData, adOpenDynamic, adLockOptimistic)
  
60      txtAcheteur.Text = rstEmploye.Fields("employe")
65      txtAcheteur.Tag = rstAchat.Fields("Acheteur")
  
70      Call rstEmploye.Close
75      Set rstEmploye = Nothing
  
80      txtRaison.Text = rstAchat.Fields("Raison")
85      txtDate.Text = rstAchat.Fields("DateAchat")
  
90      txtPrixTotal.Text = Conversion(rstAchat.Fields("PrixTotal"), MODE_ARGENT)

95      Call rstAchat.Close
100     Set rstAchat = Nothing
  
105     Call RemplirListViewAchat

110     Exit Sub

AfficherErreur:

115     woups "frmAchat", "RemplirAchat", Err, Erl
End Sub

Private Sub cmbNoAchat_KeyUp(KeyCode As Integer, Shift As Integer)
  
5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer
  
15      For iCompteur = 0 To cmbNoAchat.ListCount - 1
20        If UCase(cmbNoAchat.LIST(iCompteur)) = UCase(cmbNoAchat.Text) Then
25          cmbNoAchat.ListIndex = iCompteur
      
30          Exit For
35        End If
40      Next

45      Exit Sub

AfficherErreur:

50      woups "frmAchat", "cmbNoAchat_KeyUp", Err, Erl
End Sub

Private Sub cmdAnnuler_Click()

5       On Error GoTo AfficherErreur

10      Dim sChamps     As String
15      Dim sTable      As String
20      Dim sTablePiece As String
    
25      Screen.MousePointer = vbHourglass
  
        'Initialisation des variables booléennes
30      m_bInventaire = False
35      m_bMauvaisPrix = False
40      m_bPieceInutile = False
45      m_bRecherchePiece = False
  
        'Remet en mode inactif
50      Call AfficherControles(MODE_INACTIF)
  
55      Call OuvrirAchat(False)
    
60      Call RemplirComboAchat(m_sAncienAchat)
  
65      m_bModeAjout = False
    
70      Screen.MousePointer = vbDefault

75      Exit Sub

AfficherErreur:

80      woups "frmAchat", "cmdAnnuler_Click", Err, Erl
End Sub

Private Sub cmdAnnulerDateRequise_Click()

5       On Error GoTo AfficherErreur

10      fraDateRequise.Visible = False

15      m_bMonthViewHasFocus = False

20      Exit Sub

AfficherErreur:

25      woups "frmAchat", "cmdAnnulerDateRequise_Click", Err, Erl
End Sub

Private Sub cmdAnnulerDateRequise_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
5       On Error GoTo AfficherErreur

10      If m_bMonthViewHasFocus = True Then
15        Call cmdAnnulerDateRequise_Click
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmAchat", "cmdAnnulerDateRequise_MouseUp", Err, Erl
End Sub

Private Sub cmdAnnulerInventaire_Click()

5       On Error GoTo AfficherErreur

10      fraInventaire.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmAchat", "cmdAnnulerInventaire_Click", Err, Erl
End Sub

Private Sub cmdBonCommande_Click()

5       On Error GoTo AfficherErreur

10      Dim rstAchat As ADODB.Recordset

15      Set rstAchat = New ADODB.Recordset

20      Call rstAchat.Open("SELECT Modification, Par FROM GRB_Achat WHERE IDAchat = '" & Left$(txtNoAchat.Text, 9) & "' AND IndexAchat = " & CInt(Right$(txtNoAchat.Text, 3)), g_connData, adOpenDynamic, adLockOptimistic)

25      If rstAchat.Fields("Modification") = False Then
30        If lvwAchat.ListItems.count > 0 Then
35          If m_eCatalogue = ELECTRIQUE Then
40            Call frmChoixBonCommande.AfficherAchat(Left$(txtNoAchat.Text, 9), CInt(Right$(txtNoAchat.Text, 3)), ELECTRIQUE)
45          Else
50            Call frmChoixBonCommande.AfficherAchat(Left$(txtNoAchat.Text, 9), CInt(Right$(txtNoAchat.Text, 3)), MECANIQUE)
55          End If
60        Else
65          Call MsgBox("Il n'y a pas de pièces à commander pour cet achat!", vbOKOnly, "Erreur")
70        End If
75      Else
80        Call MsgBox("Ce projet est en modification par " & rstAchat.Fields("Par") & "!", vbOKOnly, "Erreur")
85      End If

90      Call rstAchat.Close
95      Set rstAchat = Nothing

100     Exit Sub

AfficherErreur:

105     woups "frmAchat", "cmdBonCommande_Click", Err, Erl
End Sub

Private Sub cmdDemande_Click()

5       On Error GoTo AfficherErreur

10      Dim rstAchat    As ADODB.Recordset
15      Dim sIDAchat    As String
20      Dim iIndexAchat As Integer

25      sIDAchat = Trim$(Left$(txtNoAchat.Text, InStr(1, txtNoAchat.Text, "-") + 2))
30      iIndexAchat = CInt(Trim$(Right$(txtNoAchat.Text, Len(txtNoAchat.Text) - InStrRev(txtNoAchat.Text, "-"))))

35      Set rstAchat = New ADODB.Recordset

40      Call rstAchat.Open("SELECT Modification, Par FROM GRB_Achat WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat, g_connData, adOpenDynamic, adLockOptimistic)

45      If rstAchat.Fields("Modification") = False Then
50        Call frmChoixDemande.AfficherAchat(txtNoAchat.Text, m_eCatalogue, MODE_PIECE)
55      Else
60        Call MsgBox("Cet achat est en modification par " & rstAchat.Fields("Par") & "!", vbOKOnly, "Erreur")
65      End If

70      Call rstAchat.Close
75      Set rstAchat = Nothing
        
80      Exit Sub

AfficherErreur:

85      woups "frmAchat", "cmdDemande_Click", Err, Erl
End Sub

Private Sub cmdImprimer_Click()

5       On Error GoTo AfficherErreur

10      Dim rstAchat    As ADODB.Recordset
15      Dim sIDAchat    As String
20      Dim iIndexAchat As Integer

25      sIDAchat = Trim$(Left$(txtNoAchat.Text, InStr(1, txtNoAchat.Text, "-") + 2))
30      iIndexAchat = CInt(Trim$(Right$(txtNoAchat.Text, Len(txtNoAchat.Text) - InStrRev(txtNoAchat.Text, "-"))))

35      Set rstAchat = New ADODB.Recordset

40      Call rstAchat.Open("SELECT Modification, Par FROM GRB_Achat WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat, g_connData, adOpenDynamic, adLockOptimistic)

45      If rstAchat.Fields("Modification") = False Then
50        Call frmChoixImpressionAchat.Afficher(m_eCatalogue)
55      Else
60        Call MsgBox("Cet achat est en modification par " & rstAchat.Fields("Par") & "!", vbOKOnly, "Erreur")
65      End If

70      Call rstAchat.Close
75      Set rstAchat = Nothing

80      Exit Sub

AfficherErreur:

85      woups "frmAchat", "cmdImprimer_Click", Err, Erl
End Sub

Private Sub cmdInventaire_Click()

5       On Error GoTo AfficherErreur

10      Call RemplirListViewInventaire

15      fraInventaire.Visible = True

20      Exit Sub

AfficherErreur:

25      woups "frmAchat", "cmdInventaire_Click", Err, Erl
End Sub

Private Sub cmdOKDateRequise_Click()

5       On Error GoTo AfficherErreur

10      lvwAchat.SelectedItem.SubItems(I_COL_ACHAT_DATE_REQUISE) = ConvertDate(mvwDateRequise.Value)

15      lvwAchat.SelectedItem.ListSubItems(I_COL_ACHAT_DATE_REQUISE).ForeColor = COLOR_ORANGE

20      fraDateRequise.Visible = False

25      m_bMonthViewHasFocus = False

30      Exit Sub

AfficherErreur:

35      woups "frmAchat", "cmdOKDateRequise_Click", Err, Erl
End Sub

Private Sub cmdOKDateRequise_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
5       On Error GoTo AfficherErreur

10      If m_bMonthViewHasFocus = True Then
15        Call cmdOKDateRequise_Click
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmAchat", "cmdOKDateRequise_MouseUp", Err, Erl
End Sub

Private Sub cmdOKInventaire_Click()

5       On Error GoTo AfficherErreur

        'On affiche lvwFournisseur selon la pièce choisie
        'Rempli les fournisseurs de la pièce choisie
10      m_bInventaire = True
15      m_bRecherchePiece = False
        
20      Call AfficherListeFournisseurs
  
        'Si le listview n'est pas vide
25      If lvwfournisseur.ListItems.count = 1 Then
30        If MsgBox("Il n'y a aucun fournisseur pour cette pièce!" & vbNewLine & "Voulez-vous en ajouter?", vbYesNo, "Erreur") = vbYes Then
35          Screen.MousePointer = vbHourglass
      
            'On ouvre le catalogue sur cet enregistrement
40          If m_eCatalogue = ELECTRIQUE Then
45            Call FrmCatalogueElec.AfficherForm(cmbCategorie.Text, lvwInventaire.SelectedItem.SubItems(I_COL_INV_MANUFACT), lvwInventaire.SelectedItem.Text)
50          Else
55            Call FrmCatalogueMec.AfficherForm(cmbCategorie.Text, lvwInventaire.SelectedItem.SubItems(I_COL_INV_MANUFACT), lvwInventaire.SelectedItem.Text)
60          End If
      
65          Screen.MousePointer = vbDefault
      
            'On rappelle la méthode
70          Call AfficherListeFournisseurs
75        End If
80      End If

85      fraInventaire.Visible = False

90      Exit Sub

AfficherErreur:

95      woups "frmAchat", "cmdOKInventaire_Click", Err, Erl
End Sub

Private Sub cmdRafraichir_Click()

5       On Error GoTo AfficherErreur

10      If m_sTri <> vbNullString Then
15        m_sTri = vbNullString
  
20        Call RemplirListViewPieces
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmAchat", "cmdRafraichir_Click", Err, Erl
End Sub

Private Sub cmdReset_Click()
        'Permet d'effacer le champs Modification et Par si c'est le user actuel
5       On Error GoTo AfficherErreur

10      Dim rstAchat    As ADODB.Recordset
15      Dim sIDAchat    As String
20      Dim iIndexAchat As Integer

25      If MsgBox("Êtes-vous certains de ne pas être en modification sur un autre ordinateur?", vbYesNo) = vbYes Then
30        sIDAchat = Trim$(Left$(txtNoAchat.Text, InStr(1, txtNoAchat.Text, "-") + 2))
35        iIndexAchat = CInt(Trim$(Right$(txtNoAchat.Text, Len(txtNoAchat.Text) - InStrRev(txtNoAchat.Text, "-"))))

40        Set rstAchat = New ADODB.Recordset

45        Call rstAchat.Open("SELECT Modification, Par FROM GRB_Achat WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat, g_connData, adOpenDynamic, adLockOptimistic)

50        rstAchat.Fields("Modification") = False
55        rstAchat.Fields("Par") = ""

60        Call rstAchat.Update

65        Call rstAchat.Close
70        Set rstAchat = Nothing

75        cmdReset.Visible = False
80      End If

85      Exit Sub

AfficherErreur:

90      woups "frmAchat", "cmdReset_Click", Err, Erl
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

40      woups "frmAchat", "cmdTri_Click", Err, Erl
End Sub

Private Sub lvwAchat_DblClick()

5       On Error GoTo AfficherErreur

        'Si en mode ajout ou modif
10      If m_eMode = MODE_AJOUT_MODIF Then
          'Si la liste n'est pas vide
15        If lvwAchat.ListItems.count > 0 Then
            'Si la pièce n'a pas de fournisseur
20          If Trim$(lvwAchat.SelectedItem.SubItems(I_COL_ACHAT_DISTRIB)) = vbNullString Then
25            Call ViderChamps_frs

30            cmbfrs.Locked = False

35            m_bMauvaisPrix = False

              'Rempli le combo des fournisseurs
40            Call RemplirComboFournisseur
           
              'Montre le frame
45            fraPrixPiece.Visible = True

              'Met le numéro de la pièce dans le tag
50            fraPrixPiece.Tag = lvwAchat.SelectedItem.Index
                  
              'Donne le focus au combo
55            Call cmbfrs.SetFocus
60          Else
65            'Si le listItem est orange
70            If lvwAchat.SelectedItem.ListSubItems(I_COL_ACHAT_PIECE).ForeColor = COLOR_ORANGE Then
75              If MsgBox("Voulez-vous annuler cette commande?", vbYesNo) = vbYes Then
80                Call AnnulerCommande
85              End If
90            End If
95          End If
100       End If
105     End If

110     Exit Sub

AfficherErreur:

115     woups "frmAchat", "lvwAchat_DblClick", Err, Erl
End Sub

Private Sub lvwAchat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
5       On Error GoTo AfficherErreur

10      Dim iNbreSelected  As Integer
15      Dim iIndexSelected As Integer
20      Dim iCompteur      As Integer
25      Dim bAfficherMenu  As Boolean

30      If m_eMode = MODE_AJOUT_MODIF Then
35        If Button = vbRightButton Then
40          If lvwAchat.ListItems.count > 0 Then
              'S'il y a plusieurs items de sélectionnés, c'est parce que l'utilisateur
              'a sélectionné plusieurs items
              'Donc, on ne désélectionne pas
45            For iCompteur = 1 To lvwAchat.ListItems.count
50              If lvwAchat.ListItems(iCompteur).Selected = True Then
55                iNbreSelected = iNbreSelected + 1

60                iIndexSelected = iCompteur
65              End If
70            Next

75            If iNbreSelected = 1 Then
80              lvwAchat.ListItems(iIndexSelected).Selected = False
85            End If

90            Set lvwAchat.DropHighlight = lvwAchat.HitTest(X, Y)

95            If Not lvwAchat.DropHighlight Is Nothing Then
100             If iNbreSelected = 1 Then
105               lvwAchat.DropHighlight.Selected = True

110               If lvwAchat.DropHighlight.SubItems(I_COL_ACHAT_DATE_REQUISE) = "" Then
115                 lvwAchat.DropHighlight.SubItems(I_COL_ACHAT_DATE_REQUISE) = " "
120               End If

125               If lvwAchat.DropHighlight.ListSubItems(I_COL_ACHAT_DATE_REQUISE).ForeColor = COLOR_ORANGE Then
130                 bAfficherMenu = True
135               Else
140                 bAfficherMenu = False
145               End If
150             Else
155               bAfficherMenu = False
160             End If
165           Else
170             bAfficherMenu = False
175           End If

180           If bAfficherMenu = True Then
185             Call RemplirOptionsMenuRightClick(iNbreSelected)

190             Call PopupMenu(mnuRightClick)
195           End If
200         End If
205       Else
210         If Shift <> vbCtrlMask And Shift <> vbShiftMask Then
215           Set lvwAchat.DropHighlight = Nothing
220         End If
225       End If
230     End If

235     Exit Sub

AfficherErreur:

240     woups "frmAchat", "lvwAchat_MouseDown", Err, Erl
End Sub

Private Sub RemplirOptionsMenuRightClick(ByVal iNbreSelected As Integer)

5       On Error GoTo AfficherErreur

10      Dim bDateRequise As Boolean

15      If iNbreSelected = 1 Then
          'Si c'est une sous-section
20        Select Case lvwAchat.SelectedItem.ListSubItems(I_COL_ACHAT_PIECE).ForeColor
            Case COLOR_ORANGE:
25            bDateRequise = True
30        End Select
35      End If

        'Pour empeche que tous les éléments deviennent invisible, je les mets visible au
        'début
40      mnuDateRequise.Visible = True
 
45      mnuDateRequise.Visible = bDateRequise

50      Exit Sub

AfficherErreur:

55      woups "frmAchat", "RemplirOptionsMenuRightClick", Err, Erl
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

85      woups "frmAchat", "lvwPieces_ColumnClick", Err, Erl
End Sub

Private Sub RechercherPiece(ByVal iCol As Integer, ByVal sTexte As String)

5       On Error GoTo AfficherErreur

10      Dim sValeur     As String
15      Dim rstcat      As ADODB.Recordset
20      Dim iCompteur   As Integer
25      Dim bTrouverLvw As Boolean
30      Dim bTrouverRst As Boolean
35      Dim iIndexCat   As Integer
40      Dim sChamps     As String
45      Dim sCategorie  As String
    
50      For iCompteur = 1 To lvwPieces.ListItems.count
55        If iCol > 0 Then
60          sValeur = lvwPieces.ListItems(iCompteur).SubItems(iCol)
65        Else
70          sValeur = lvwPieces.ListItems(iCompteur).Text
75        End If
    
80        sValeur = UCase(sValeur)
85        sTexte = UCase(sTexte)
        
90        If InStr(1, sValeur, sTexte) > 0 Then
95          lvwPieces.ListItems(iCompteur).Selected = True
        
100         Call lvwPieces.SelectedItem.EnsureVisible
        
105         bTrouverLvw = True
110       End If
    
115       If bTrouverLvw = True Then
120         Exit For
125       End If
130     Next
  
135     Select Case iCol
          Case I_COL_PIECES_PIECE_GRB: sChamps = "PIECE_GRB"
140       Case I_COL_PIECES_NO_ITEM:   sChamps = "PIECE"
145       Case I_COL_PIECES_MANUFACT:  sChamps = "FABRICANT"
150       Case I_COL_PIECES_DESCR_FR:  sChamps = "DESC_FR"
155       Case I_COL_PIECES_DESCR_EN:  sChamps = "DESC_EN"
160     End Select
  
165     iIndexCat = cmbCategorie.ListIndex
  
170     If bTrouverLvw = False Then
175       Set rstcat = New ADODB.Recordset

180       For iCompteur = iIndexCat + 1 To cmbCategorie.ListCount - 1
185         sCategorie = Replace(cmbCategorie.LIST(iCompteur), "'", "''")
              
190         If m_eCatalogue = ELECTRIQUE Then
195           Call rstcat.Open("SELECT * FROM GRB_CatalogueElec WHERE INSTR(1, " & sChamps & ", '" & sTexte & "') > 0 AND CATEGORIE = '" & sCategorie & "'", g_connData, adOpenDynamic, adLockOptimistic)
200         Else
205           Call rstcat.Open("SELECT * FROM GRB_CatalogueMec WHERE INSTR(1, " & sChamps & ", '" & sTexte & "') > 0 AND CATEGORIE = '" & sCategorie & "'", g_connData, adOpenDynamic, adLockOptimistic)
210         End If
          
215         If Not rstcat.EOF Then
220           bTrouverRst = True
        
225           cmbCategorie.ListIndex = iCompteur
        
230           Call RechercherPiece(iCol, sTexte)
       
235           Exit For
240         End If
      
245         Call rstcat.Close
250       Next
    
255       If bTrouverRst = False Then
260         For iCompteur = 0 To iIndexCat - 1
265           sCategorie = Replace(cmbCategorie.LIST(iCompteur), "'", "''")
             
270           If m_eCatalogue = ELECTRIQUE Then
275             Call rstcat.Open("SELECT * FROM GRB_CatalogueElec WHERE INSTR(1, " & sChamps & ", '" & sTexte & "') > 0 AND CATEGORIE = '" & sCategorie & "'", g_connData, adOpenDynamic, adLockOptimistic)
280           Else
285             Call rstcat.Open("SELECT * FROM GRB_CatalogueMec WHERE INSTR(1, " & sChamps & ", '" & sTexte & "') > 0 AND CATEGORIE = '" & sCategorie & "'", g_connData, adOpenDynamic, adLockOptimistic)
290           End If
          
295           If Not rstcat.EOF Then
300             bTrouverRst = True
        
305             cmbCategorie.ListIndex = iCompteur
        
310             Call RechercherPiece(iCol, sTexte)
       
315             Exit For
320           End If
      
325           Call rstcat.Close
330         Next
      
335         If bTrouverRst = False Then
340           Call MsgBox("Aucun enregistrements trouvés!", vbOKOnly, "Erreur")
345         End If
350       End If

355       Set rstcat = Nothing
360     End If

365     Exit Sub

AfficherErreur:

370     woups "frmAchat", "RechercherPiece", Err, Erl
End Sub

Private Sub lvwPieces_KeyDown(KeyCode As Integer, Shift As Integer)

5       On Error GoTo AfficherErreur

10      Dim sTexte As String
  
15      If Shift = vbCtrlMask Then
20        If KeyCode = vbKeyF Then
25          sTexte = InputBox("Quel est le texte à rechercher?")
  
30          If Trim$(sTexte) <> vbNullString Then
35            If Len(Trim$(sTexte)) >= 3 Then
40              Call RechercherPiece(I_COL_PIECES_NO_ITEM, sTexte)
45            Else
50              Call MsgBox("Il faut un minimum de 3 caractères pour rechercher!", vbOKOnly, "Erreur")
55            End If
60          End If
65        End If
70      End If

75      Exit Sub

AfficherErreur:

80      woups "frmAchat", "lvwPieces_KeyDown", Err, Erl
End Sub

Private Sub cmdEnregistrer_Click()

5       On Error GoTo AfficherErreur

10      Dim objControl  As Control
15      Dim iIndexAchat As Integer
  
        'Vérification des textbox
20      Screen.MousePointer = vbHourglass
        
25      For Each objControl In Me
30        If TypeOf objControl Is TextBox Then
35          If objControl.Visible = True Then
40            If Trim$(objControl.Text) = vbNullString Then
45              Screen.MousePointer = vbDefault

50              Call MsgBox("Vous devez remplir tous les champs!", vbOKOnly, "Erreur")
        
55              Exit Sub
60            End If
65          End If
70        End If
75      Next
      
80      If BackupPieces(txtNoAchat.Text) = False Then
85        Screen.MousePointer = vbDefault

90        If MsgBox("Une erreur est survenue lors de la copie de sauvegarde de l'achat en cours!" & vbNewLine & _
                    vbNewLine & _
                    "Voulez-vous continuer?", vbYesNo) = vbNo Then
95          Exit Sub
100       Else
105         Screen.MousePointer = vbHourglass
110       End If
115     End If
        
        'Enregistre l'achat
120     Call EnregistrerAchat(txtNoAchat.Text)
  
        'Initialisation des variables booléennes
125     m_bInventaire = False
130     m_bMauvaisPrix = False
135     m_bPieceInutile = False
140     m_bRecherchePiece = False
  
145     Call OuvrirAchat(False)
  
        'Remet en mode inactif
150     Call AfficherControles(MODE_INACTIF)
  
        'Affiche l'achat actuel
155     If Len(txtNoAchat.Text) = 9 Then
160       iIndexAchat = TrouverNouvelIndex
165     End If

170     If iIndexAchat > 0 Then
175       Call AfficherAchat(txtNoAchat.Text & "-" & Right$("00" & iIndexAchat, 3))
180     Else
185       Call AfficherAchat(txtNoAchat.Text)
190     End If
  
195     Screen.MousePointer = vbDefault

200     Exit Sub

AfficherErreur:

205     woups "frmAchat", "cmdEnregistrer_Click", Err, Erl
End Sub

Private Function BackupPieces(ByVal sNoAchat As String) As Boolean

5       On Error GoTo AfficherErreur

10      Dim rstAchat       As ADODB.Recordset
15      Dim rstAchatBackup As ADODB.Recordset
20      Dim sDateCopie     As String
25      Dim sIDAchat       As String
30      Dim iIndexAchat    As Integer

35      If m_bModeAjout = False Then
40        sIDAchat = Left$(sNoAchat, 9)

45        iIndexAchat = Right$(sNoAchat, 3)
50      Else
55        BackupPieces = True

60        Exit Function
65      End If

70      Set rstAchat = New ADODB.Recordset
75      Set rstAchatBackup = New ADODB.Recordset

80      Call rstAchat.Open("SELECT * FROM GRB_Achat_Pieces WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat, g_connData, adOpenForwardOnly, adLockReadOnly)

85      Call rstAchatBackup.Open("SELECT * FROM GRB_Achat_Pieces_Tampon", g_connData, adOpenDynamic, adLockOptimistic)

90      sDateCopie = ConvertDate(Date) & " " & Time

95      Do While Not rstAchat.EOF
100       Call rstAchatBackup.AddNew

105       rstAchatBackup.Fields("DateCopie") = sDateCopie

110       rstAchatBackup.Fields("IDAchat") = rstAchat.Fields("IDAchat")
115       rstAchatBackup.Fields("IndexAchat") = rstAchat.Fields("IndexAchat")

120       rstAchatBackup.Fields("Initiales") = g_sInitiale
125       rstAchatBackup.Fields("PIECE") = rstAchat.Fields("PIECE")
130       rstAchatBackup.Fields("NuméroLigne") = rstAchat.Fields("NuméroLigne")
135       rstAchatBackup.Fields("Qté") = rstAchat.Fields("Qté")
140       rstAchatBackup.Fields("Desc_FR") = rstAchat.Fields("Desc_FR")
145       rstAchatBackup.Fields("Desc_EN") = rstAchat.Fields("Desc_EN")
150       rstAchatBackup.Fields("Manufact") = rstAchat.Fields("Manufact")
155       rstAchatBackup.Fields("Prix_list") = rstAchat.Fields("Prix_list")
160       rstAchatBackup.Fields("Escompte") = rstAchat.Fields("Escompte")
165       rstAchatBackup.Fields("Prix_net") = rstAchat.Fields("Prix_net")
170       rstAchatBackup.Fields("IDFRS") = rstAchat.Fields("IDFRS")
175       rstAchatBackup.Fields("Prix_total") = rstAchat.Fields("Prix_total")
180       rstAchatBackup.Fields("Type") = rstAchat.Fields("Type")
185       rstAchatBackup.Fields("Commandé") = rstAchat.Fields("Commandé")
190       rstAchatBackup.Fields("Retour") = rstAchat.Fields("Retour")
195       rstAchatBackup.Fields("NoRetour") = rstAchat.Fields("NoRetour")
200       rstAchatBackup.Fields("Recu") = rstAchat.Fields("Recu")
205       rstAchatBackup.Fields("DateRéception") = rstAchat.Fields("DateRéception")
210       rstAchatBackup.Fields("QuantitéRecue") = rstAchat.Fields("QuantitéRecue")
215       rstAchatBackup.Fields("DateCommande") = rstAchat.Fields("DateCommande")
220       rstAchatBackup.Fields("DateRequise") = rstAchat.Fields("DateRequise")
225       rstAchatBackup.Fields("Inutile") = rstAchat.Fields("Inutile")
230       rstAchatBackup.Fields("CommandeAnnulée") = rstAchat.Fields("CommandeAnnulée")
235       rstAchatBackup.Fields("DateRetour") = rstAchat.Fields("DateRetour")
240       rstAchatBackup.Fields("PrixOrigine") = rstAchat.Fields("PrixOrigine")
245       rstAchatBackup.Fields("Devise") = rstAchat.Fields("Devise")

250       Call rstAchatBackup.Update

255       Call rstAchat.MoveNext
260     Loop

265     Call rstAchat.Close
270     Set rstAchat = Nothing

275     Call rstAchatBackup.Close
280     Set rstAchatBackup = Nothing

285     BackupPieces = True

290     Exit Function

AfficherErreur:

295     woups "frmAchat", "BackupPieces", Err, Erl
End Function

Private Function TrouverNouvelIndex() As Integer

5       On Error GoTo AfficherErreur

10      Dim rstMax As ADODB.Recordset
15      Dim iIndex As Integer
  
20      Set rstMax = New ADODB.Recordset
  
25      Call rstMax.Open("SELECT MAX(IndexAchat) AS MaxIndex FROM GRB_Achat WHERE IDAchat = '" & txtNoAchat.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
  
30      iIndex = rstMax.Fields("MaxIndex")
  
35      Call rstMax.Close
40      Set rstMax = Nothing
  
45      TrouverNouvelIndex = iIndex

50      Exit Function

AfficherErreur:

55      woups "frmAchat", "TrouverNouvelIndex", Err, Erl
End Function

Private Sub EnregistrerAchat(ByVal sNoAchat As String)

5       On Error GoTo AfficherErreur

10      Dim rstAchat     As ADODB.Recordset
15      Dim rstPiece     As ADODB.Recordset
20      Dim rstMax       As ADODB.Recordset
25      Dim itmPiece     As ListItem
30      Dim dblPrixTotal As Double
35      Dim sIDAchat     As String
40      Dim iIndexAchat  As Integer
45      Dim iCompteur    As Integer
        Dim testgll As String
50      sIDAchat = Left$(sNoAchat, 9)
    
55      Set rstAchat = New ADODB.Recordset
    
        'Si c'est un ajout
60      If m_bModeAjout = True Then
          'On ouvre le recordset
65        Call rstAchat.Open("SELECT * FROM GRB_Achat", g_connData, adOpenDynamic, adLockOptimistic)
    
          'Ajouter une nouvelle achat
70        Call rstAchat.AddNew
         
75        m_bModeAjout = False
80      Else
85        iIndexAchat = Right$(sNoAchat, 3)
  
90       Call rstAchat.Open("SELECT * FROM GRB_Achat WHERE IDAchat" & " = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat, g_connData, adOpenDynamic, adLockOptimistic)
      
          'Si c'est une modification, il faut effacer les pieces et remplir les nouvelles
95        Call g_connData.Execute("DELETE * FROM GRB_Achat_Pieces WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat)
100     End If
  
        'Enregistrement de l'achat
  
        'IDAchat
105     rstAchat.Fields("IDAchat") = sIDAchat
  
        'IndexAchat
110     If iIndexAchat = 0 Then
115       Set rstMax = New ADODB.Recordset

          'Pour avoir le dernier index
120       Call rstMax.Open("SELECT MAX(IndexAchat) As MaxAchat FROM GRB_Achat WHERE IDAchat = '" & sNoAchat & "'", g_connData, adOpenDynamic, adLockOptimistic)
    
125       If Not IsNull(rstMax("MaxAchat")) Then
130         rstAchat.Fields("IndexAchat") = rstMax("MaxAchat") + 1
135       Else
140         rstAchat.Fields("IndexAchat") = 1
145       End If
    
150       Call rstMax.Close
155       Set rstMax = Nothing
160     Else
165       rstAchat.Fields("IndexAchat") = iIndexAchat
170     End If
  
175     rstAchat.Fields("Raison") = txtRaison.Text
180     rstAchat.Fields("DateAchat") = txtDate.Text
185     rstAchat.Fields("Acheteur") = txtAcheteur.Tag
      
190     If m_eCatalogue = ELECTRIQUE Then
195       rstAchat.Fields("Type") = "E"
200     Else
205       rstAchat.Fields("Type") = "M"
210     End If

215     Set rstPiece = New ADODB.Recordset

220     Call rstPiece.Open("SELECT * FROM GRB_Achat_Pieces", g_connData, adOpenDynamic, adLockOptimistic)
    
        'Enregistrement des pièces
225     For iCompteur = 1 To lvwAchat.ListItems.count
230       Set itmPiece = lvwAchat.ListItems(iCompteur)
      
235       Call rstPiece.AddNew
         
240       rstPiece.Fields("IDAchat") = rstAchat.Fields("IDAchat")
245       rstPiece.Fields("IndexAchat") = rstAchat.Fields("IndexAchat")
        
    
250       rstPiece.Fields("PIECE") = itmPiece.SubItems(I_COL_ACHAT_PIECE)
255       rstPiece.Fields("NuméroLigne") = iCompteur
260       rstPiece.Fields("Qté") = itmPiece.Text
265       rstPiece.Fields("Desc_FR") = itmPiece.SubItems(I_COL_ACHAT_DESCR)
270       rstPiece.Fields("Desc_EN") = itmPiece.ListSubItems(I_COL_ACHAT_DESCR).Tag
275       rstPiece.Fields("Manufact") = itmPiece.SubItems(I_COL_ACHAT_MANUFACT)
280       rstPiece.Fields("Prix_list") = Conversion(itmPiece.SubItems(I_COL_ACHAT_PRIX_LIST), MODE_PAS_FORMAT, 4)

285       If itmPiece.SubItems(I_COL_ACHAT_PRIX_LIST) <> vbNullString Then
290         If itmPiece.ListSubItems(I_COL_ACHAT_PRIX_LIST).Tag <> vbNullString Then
295           rstPiece.Fields("PrixOrigine") = Replace(Round(CDbl(Replace(itmPiece.ListSubItems(I_COL_ACHAT_PRIX_LIST).Tag, ".", ",")), 4), ".", ",")
300         Else
305           rstPiece.Fields("PrixOrigine") = "0"
310         End If
315       Else
320         rstPiece.Fields("PrixOrigine") = "0"
325       End If
    
330       If itmPiece.SubItems(I_COL_ACHAT_TOTAL) <> "" Then
335         rstPiece.Fields("Devise") = itmPiece.ListSubItems(I_COL_ACHAT_TOTAL).Tag
340       Else
345         rstPiece.Fields("Devise") = ""
350       End If

355       If Trim$(itmPiece.SubItems(I_COL_ACHAT_ESCOMPTE)) <> "" Then
360         rstPiece.Fields("Escompte") = Conversion(Replace(itmPiece.SubItems(I_COL_ACHAT_ESCOMPTE), "%", "") / 100, MODE_PAS_FORMAT)
365       Else
370         rstPiece.Fields("Escompte") = ""
375       End If

380       rstPiece.Fields("Prix_net") = Conversion(itmPiece.SubItems(I_COL_ACHAT_PRIX_NET), MODE_PAS_FORMAT, 4)
385       rstPiece.Fields("DateRéception") = itmPiece.Tag
390       rstPiece.Fields("NoRetour") = itmPiece.ListSubItems(I_COL_ACHAT_MANUFACT).Tag

395       If itmPiece.ListSubItems(I_COL_ACHAT_DISTRIB).Tag <> "" Then
400         rstPiece.Fields("IDFRS") = itmPiece.ListSubItems(I_COL_ACHAT_DISTRIB).Tag
405       Else
410         rstPiece.Fields("IDFRS") = 0
415       End If

420       If itmPiece.ListSubItems(I_COL_ACHAT_PIECE).ForeColor = COLOR_ORANGE Then
425         rstPiece.Fields("Commandé") = True
430       Else
435         rstPiece.Fields("Commandé") = False
440       End If

445       If itmPiece.ListSubItems(I_COL_ACHAT_PIECE).ForeColor = COLOR_BLEU Or itmPiece.ListSubItems(I_COL_ACHAT_PIECE).ForeColor = COLOR_GRIS Then
450         rstPiece.Fields("Recu") = True
455       Else
460         rstPiece.Fields("Recu") = False
465       End If

470       If itmPiece.ListSubItems(I_COL_ACHAT_PIECE).ForeColor = COLOR_ROUGE Then
475         rstPiece.Fields("Retour") = True
480       Else
485         rstPiece.Fields("Retour") = False
490       End If

495       If itmPiece.ListSubItems(I_COL_ACHAT_PIECE).ForeColor = COLOR_BRUN Then
500         rstPiece.Fields("Inutile") = True
505       Else
510         rstPiece.Fields("Inutile") = False
515       End If

520       If itmPiece.ListSubItems(I_COL_ACHAT_PIECE).ForeColor = COLOR_VERT_FORET Then
525         rstPiece.Fields("CommandeAnnulée") = True
530       Else
535         rstPiece.Fields("CommandeAnnulée") = False
540       End If

545       rstPiece.Fields("Prix_Total") = Conversion(itmPiece.SubItems(I_COL_ACHAT_TOTAL), MODE_PAS_FORMAT)

550       If itmPiece.SubItems(I_COL_ACHAT_DATE_COMMANDE) <> "" Then
555         rstPiece.Fields("DateCommande") = itmPiece.SubItems(I_COL_ACHAT_DATE_COMMANDE)
560       Else
565         rstPiece.Fields("DateCommande") = ""
570       End If
              
575       If itmPiece.SubItems(I_COL_ACHAT_DATE_REQUISE) <> "" Then
580         rstPiece.Fields("DateRequise") = itmPiece.SubItems(I_COL_ACHAT_DATE_REQUISE)
585       Else
590         rstPiece.Fields("DateRequise") = ""
595       End If
              
600       Call rstPiece.Update
        
605       If rstPiece.Fields("Prix_Total") <> vbNullString Then
610         dblPrixTotal = dblPrixTotal + rstPiece.Fields("Prix_Total")
615       End If
620     Next

625     rstAchat.Fields("PrixTotal") = CStr(dblPrixTotal)
  
630     Call rstAchat.Update
  
635     Call rstAchat.Close
640     Set rstAchat = Nothing
  
645     Call rstPiece.Close
650     Set rstPiece = Nothing

655     Exit Sub

AfficherErreur:

660     woups "frmAchat", "EnregistrerAchat", Err, Erl

665     If Erl >= 230 And Erl <= 615 Then
670       Call MsgBox("La pièce " & itmPiece.SubItems(I_COL_ACHAT_PIECE) & " risque de contenir des erreurs." & vbNewLine & _
                "Il se peut qu'elle ne soit plus présente dans la liste.")
675     End If
  
680     Resume Next
End Sub

Private Sub AfficherAchat(ByVal sNoAchat As String)

5       On Error GoTo AfficherErreur

        'Remet en mode affichage le projet ou l'achat voulue
10      m_bModeAffichage = True
    
        'Vide les champs
15      Call ViderChamps
  
        'Rempli le combo
20      Call RemplirComboAchat(sNoAchat)

25      Exit Sub

AfficherErreur:

30      woups "frmAchat", "AfficherAchat", Err, Erl
End Sub

Private Sub Cmdfermer_Click()

5       On Error GoTo AfficherErreur

10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmAchat", "cmdFermer_Click", Err, Erl
End Sub

Private Sub Cmdajouter_Click()

5       On Error GoTo AfficherErreur

        'Ajoute une achat
10      Dim rstEmploye As ADODB.Recordset
        
        'Initialisation des variables booléennes
15      m_bInventaire = False
20      m_bMauvaisPrix = False
25      m_bPieceInutile = False
30      m_bRecherchePiece = False
        
35      Call frmAjoutAchat.Afficher(m_eCatalogue)
  
40      If m_bAnnuler = False Then
45        If m_sNoAchat <> vbNullString Then
            'Vide les champs
50          Call ViderChamps
    
55          txtAcheteur.Text = g_sEmploye

60          Set rstEmploye = New ADODB.Recordset
      
65          Call rstEmploye.Open("SELECT NoEmploye FROM GRB_Employés WHERE Initiale = '" & g_sInitiale & "'", g_connData, adOpenDynamic, adLockOptimistic)
      
70          txtAcheteur.Tag = rstEmploye.Fields("NoEmploye")
     
75          Call rstEmploye.Close
80          Set rstEmploye = Nothing
          
85          txtDate.Text = ConvertDate(Date)
      
            'Pour pouvoir réafficher l'élément affiché avant l'ajout au cas où la personne
            'annule l'ajout
90          m_sAncienAchat = txtNoAchat.Text
    
            'Affiche le nouveau numéro
95          txtNoAchat.Text = m_sNoAchat
             
100         m_bModeAjout = True
105         m_bModeAffichage = False
                       
            'Met le form en mode ajout/modif
110         Call AfficherControles(MODE_AJOUT_MODIF)
115       End If
120     End If

125     Exit Sub

AfficherErreur:

130     woups "frmAchat", "cmdAjouter_Click", Err, Erl
End Sub

Private Sub RemplirComboCategorie()

5       On Error GoTo AfficherErreur

        'Remplir le combo des tables (Pièces)
10      Dim rstCategorie As ADODB.Recordset
15      Dim sNomTable    As String
  
        'Il faut vider le combo avant de le remplir
20      Call cmbCategorie.Clear
  
25      Set rstCategorie = New ADODB.Recordset

        'On rempli le recordset avec le nom de chaque catégorie
30      If m_eCatalogue = ELECTRIQUE Then
35        Call rstCategorie.Open("SELECT DISTINCT CATEGORIE FROM GRB_CatalogueElec", g_connData, adOpenDynamic, adLockOptimistic)
40      Else
45        Call rstCategorie.Open("SELECT DISTINCT CATEGORIE FROM GRB_CatalogueMec", g_connData, adOpenDynamic, adLockOptimistic)
50      End If
  
        'Tant que ce n'est pas la fin des enregistrements
55      Do While Not rstCategorie.EOF
60        If Not IsNull(rstCategorie.Fields("CATEGORIE")) Then
            'On ajoute le nom de la catégorie dans le combo
65          Call cmbCategorie.AddItem(rstCategorie.Fields("CATEGORIE"))
70        End If
      
75        Call rstCategorie.MoveNext
80      Loop
  
85      Call rstCategorie.Close
90      Set rstCategorie = Nothing

95      If cmbCategorie.ListCount > 0 Then
100       cmbCategorie.ListIndex = 0
105     End If

110     Exit Sub

AfficherErreur:

115     woups "frmAchat", "RemplirComboCategorie", Err, Erl
End Sub

Private Sub cmdModifier_Click()

5       On Error GoTo AfficherErreur

10      Dim rstAchat    As ADODB.Recordset
15      Dim sIDAchat    As String
20      Dim iIndexAchat As Integer

        'Modifier un achat
        
        'Initialisation des variables booléennes
25      m_bInventaire = False
30      m_bMauvaisPrix = False
35      m_bPieceInutile = False
40      m_bRecherchePiece = False
        
45      If cmbNoAchat.ListIndex > -1 Then
50        sIDAchat = Trim$(Left$(txtNoAchat.Text, InStr(1, txtNoAchat.Text, "-") + 2))
55        iIndexAchat = CInt(Trim$(Right$(txtNoAchat.Text, Len(txtNoAchat.Text) - InStrRev(txtNoAchat.Text, "-"))))

60        Set rstAchat = New ADODB.Recordset

65        Call rstAchat.Open("SELECT Modification, Par FROM GRB_Achat WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat, g_connData, adOpenDynamic, adLockOptimistic)

70        If rstAchat.Fields("Modification") = False Then
75          Call rstAchat.Close
80          Set rstAchat = Nothing

85          Screen.MousePointer = vbHourglass
   
            'Pour pouvoir afficher le dernier enregistrement affiché quand la personne va
            'enregistrer ou annuler
90          m_sAncienAchat = txtNoAchat.Text
    
95          m_bModeAjout = False
100         m_bModeAffichage = False

105         Call OuvrirAchat(True)
  
110         Call AfficherControles(MODE_AJOUT_MODIF)
          
115         Screen.MousePointer = vbDefault
120       Else
125         Call MsgBox("Cet achat est en modification par " & rstAchat.Fields("Par") & "!", vbOKOnly, "Erreur")

130         Call rstAchat.Close
135         Set rstAchat = Nothing
140       End If
145     End If

150     Exit Sub

AfficherErreur:

155     woups "frmAchat", "cmdModifier_Click", Err, Erl
End Sub

Private Sub cmdsupprimer_Click()

5       On Error GoTo AfficherErreur

10      Dim iReponse    As Integer
15      Dim sIDAchat    As String
20      Dim iIndexAchat As Integer
25      Dim rstAchat    As ADODB.Recordset
 
30      If cmbNoAchat.ListCount > 0 Then
35        sIDAchat = Trim$(Left$(txtNoAchat.Text, InStr(1, txtNoAchat.Text, "-") + 2))
40        iIndexAchat = CInt(Trim$(Right$(txtNoAchat.Text, Len(txtNoAchat.Text) - InStrRev(txtNoAchat.Text, "-"))))

45        Set rstAchat = New ADODB.Recordset

50        Call rstAchat.Open("SELECT Modification, Par FROM GRB_Achat WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat, g_connData, adOpenDynamic, adLockOptimistic)

55        If rstAchat.Fields("Modification") = False Then
60          Call rstAchat.Close
65          Set rstAchat = Nothing

            'Valider le choix
70          iReponse = MsgBox("Voulez-vous vraiment effacer l'achat " & txtNoAchat.Text & "?", vbYesNo)
    
            'Si il veut vraiment effacer
75          If iReponse = vbYes Then
              'Efface les pièces
80            sIDAchat = Left$(txtNoAchat.Text, 9)
            
85            iIndexAchat = Right$(txtNoAchat.Text, 3)
      
              'Efface les pièces
90            Call g_connData.Execute("DELETE * FROM GRB_Achat_Pieces WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat)
        
              'Efface l'achat
95            Call g_connData.Execute("DELETE * FROM GRB_Achat WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat)
              
100           'Affiche la premiere achat
105           Call RemplirComboAchat(vbNullString)
110         End If
115       Else
120         Call MsgBox("Cet achat est en modification par " & rstAchat.Fields("Par") & "!", vbOKOnly, "Erreur")

125         Call rstAchat.Close
130         Set rstAchat = Nothing
135       End If
140     End If

145     Exit Sub

AfficherErreur:

150     woups "frmAchat", "cmdSupprimer_Click", Err, Erl
End Sub

Private Sub ViderChamps()

5       On Error GoTo AfficherErreur

        'Méthode qui initialise les champs
10      txtPrixTotal.Text = 0
15      txtDate.Text = vbNullString
20      txtRaison.Text = vbNullString
25      txtAcheteur.Text = vbNullString
       
30      Call lvwAchat.ListItems.Clear

35      Exit Sub

AfficherErreur:

40      woups "frmAchat", "ViderChamps", Err, Erl
End Sub

Private Sub RemplirComboAchat(ByVal sNoAchat As String)

5       On Error GoTo AfficherErreur

        'Rempli le combo des achats
10      Dim rstAchat  As ADODB.Recordset
15      Dim sType     As String
20      Dim iCompteur As Integer
  
        'Il faut vider le combo avant de le remplir
25      Call cmbNoAchat.Clear
  
30      If m_eCatalogue = ELECTRIQUE Then
35        sType = "E"
40      Else
45        sType = "M"
50      End If
  
55      Set rstAchat = New ADODB.Recordset
  
60      Call rstAchat.Open("SELECT * FROM GRB_Achat WHERE Type = '" & sType & "' ORDER BY IDAchat DESC, IndexAchat DESC", g_connData, adOpenDynamic, adLockOptimistic)
  
        'Tant que ce n'est pas la fin des enregistrements
65      Do While Not rstAchat.EOF
          'On met le numéro de l'achat dans le combo des achats
70        Call cmbNoAchat.AddItem(rstAchat.Fields("IDAchat") & "-" & Right$("00" & rstAchat("IndexAchat"), 3))
      
75        Call rstAchat.MoveNext
80      Loop
      
85      Call rstAchat.Close
90      Set rstAchat = Nothing
  
        'Si le combo n'est pas vide, on sélectionne l'item voulu ou le premier
95      If cmbNoAchat.ListCount > 0 Then
100       If sNoAchat <> vbNullString Then
105         For iCompteur = 0 To cmbNoAchat.ListCount - 1
110           If cmbNoAchat.LIST(iCompteur) = sNoAchat Then
115             cmbNoAchat.ListIndex = iCompteur

120             Exit For
125           End If
130         Next
135       Else
140         cmbNoAchat.ListIndex = 0
145       End If
150     Else
155       Call ViderChamps
160     End If

165     Exit Sub

AfficherErreur:

170     woups "frmAchat", "RemplirComboAchat", Err, Erl
End Sub

Private Sub CalculerPrixReel(ByVal sNoItem As String)

5       On Error GoTo AfficherErreur

10      Dim rstPieceFRS As ADODB.Recordset
15      Dim rstConfig   As ADODB.Recordset
20      Dim sPrixCalcul As String
25      Dim sTauxUSA    As String
30      Dim sTauxSPA    As String
35      Dim sType       As String

40      If m_eCatalogue = ELECTRIQUE Then
45        sType = "E"
50      Else
55        sType = "M"
60      End If

65      Set rstConfig = New ADODB.Recordset

70      Call rstConfig.Open("SELECT TauxAmericain, TauxEspagnol FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

75      sTauxUSA = rstConfig.Fields("TauxAmericain")
80      sTauxSPA = rstConfig.Fields("TauxEspagnol")

85      Call rstConfig.Close
90      Set rstConfig = Nothing
  
95      Set rstPieceFRS = New ADODB.Recordset

100     rstPieceFRS.CursorLocation = adUseServer
  
105     Call rstPieceFRS.Open("SELECT PrixReel, PRIX_NET, PRIX_SP, DeviseMonétaire FROM GRB_PiecesFRS WHERE PIECE = '" & Replace(sNoItem, "'", "''") & "' AND Type = '" & sType & "'", g_connData, adOpenDynamic, adLockOptimistic)
    
110     Do While Not rstPieceFRS.EOF
115       If rstPieceFRS.Fields("PRIX_NET") <> vbNullString Then
120         sPrixCalcul = rstPieceFRS.Fields("PRIX_NET")
125       Else
130         If rstPieceFRS.Fields("PRIX_SP") <> vbNullString Then
135           sPrixCalcul = rstPieceFRS.Fields("PRIX_SP")
140         End If
145       End If
      
150       sPrixCalcul = Replace(sPrixCalcul, ".", ",")
      
155       If rstPieceFRS.Fields("DeviseMonétaire") = "USA" Then
160         rstPieceFRS.Fields("PrixReel") = Conversion(CStr(Round(CDbl(sPrixCalcul) / CDbl(sTauxUSA), 4)), MODE_DECIMAL, 4)
165       Else
170         If rstPieceFRS.Fields("DeviseMonétaire") = "SPA" Then
175           rstPieceFRS.Fields("PrixReel") = Conversion(CStr(Round(CDbl(sPrixCalcul) / CDbl(sTauxSPA), 4)), MODE_DECIMAL, 4)
180         Else
185            rstPieceFRS.Fields("PrixReel") = Conversion(sPrixCalcul, MODE_DECIMAL, 4)
190         End If
195       End If
      
200       Call rstPieceFRS.Update
    
205       Call rstPieceFRS.MoveNext
210     Loop
    
215     Call rstPieceFRS.Close
220     Set rstPieceFRS = Nothing

225     Exit Sub

AfficherErreur:

230     woups "frmAchat", "CalculerPrixReel", Err, Erl
End Sub

Private Sub RemplirListViewFournisseur()

5       On Error GoTo AfficherErreur

        'Rempli le listview des distributeur pour une pièce choisie
10      Dim rstPieceFRS As ADODB.Recordset
15      Dim rstContact  As ADODB.Recordset
20      Dim rstFRS      As ADODB.Recordset
25      Dim iCompteur   As Integer
30      Dim itmFRS      As ListItem
35      Dim sDevise     As String
40      Dim iNoClient   As Integer
45      Dim lColor      As Long
50      Dim sType       As String
 
55      If m_eCatalogue = ELECTRIQUE Then
60        sType = "E"
65      Else
70        sType = "M"
75      End If
 
        'vide le lister
80      Call lvwfournisseur.ListItems.Clear
      
85      Set rstPieceFRS = New ADODB.Recordset
      
90      If m_bPieceInutile = True Then
95        Call CalculerPrixReel(Trim$(lvwAchat.SelectedItem.SubItems(I_COL_ACHAT_PIECE)))

100       Call rstPieceFRS.Open("SELECT GRB_PiecesFRS.*, GRB_Fournisseur.NomFournisseur FROM GRB_PiecesFRS INNER JOIN GRB_Fournisseur ON GRB_PiecesFRS.IDFRS = GRB_Fournisseur.IDFRS WHERE PIECE = '" & Trim$(Replace(lvwAchat.SelectedItem.SubItems(I_COL_ACHAT_PIECE), "'", "''")) & "' AND Type = '" & sType & "' ORDER BY PrixReel", g_connData, adOpenDynamic, adLockOptimistic)
105     Else
110       If m_bInventaire = True Then
115         Call CalculerPrixReel(Trim$(lvwInventaire.SelectedItem.Text))

120         Call rstPieceFRS.Open("SELECT GRB_PiecesFRS.*, GRB_Fournisseur.NomFournisseur FROM GRB_PiecesFRS INNER JOIN GRB_Fournisseur ON GRB_PiecesFRS.IDFRS = GRB_Fournisseur.IDFRS WHERE PIECE = '" & Trim$(Replace(lvwInventaire.SelectedItem.Text, "'", "''")) & "' AND Type = '" & sType & "' ORDER BY PrixReel", g_connData, adOpenDynamic, adLockOptimistic)
125       Else
130         If m_bRecherchePiece = True Then
135           Call CalculerPrixReel(Trim$(lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM)))

140           Call rstPieceFRS.Open("SELECT GRB_PiecesFRS.*, GRB_Fournisseur.NomFournisseur FROM GRB_PiecesFRS INNER JOIN GRB_Fournisseur ON GRB_PiecesFRS.IDFRS = GRB_Fournisseur.IDFRS WHERE PIECE = '" & Trim$(Replace(lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM), "'", "''")) & "' AND Type = '" & sType & "' ORDER BY PrixReel", g_connData, adOpenDynamic, adLockOptimistic)
145         Else
150           Call CalculerPrixReel(Trim$(lvwPieces.SelectedItem.SubItems(I_COL_PIECES_NO_ITEM)))

155           Call rstPieceFRS.Open("SELECT GRB_PiecesFRS.*, GRB_Fournisseur.NomFournisseur FROM GRB_PiecesFRS INNER JOIN GRB_Fournisseur ON GRB_PiecesFRS.IDFRS = GRB_Fournisseur.IDFRS WHERE PIECE = '" & Trim$(Replace(lvwPieces.SelectedItem.SubItems(I_COL_PIECES_NO_ITEM), "'", "''")) & "' AND Type = '" & sType & "' ORDER BY PrixReel", g_connData, adOpenDynamic, adLockOptimistic)
160         End If
165       End If
170     End If

175     Set rstFRS = New ADODB.Recordset

180     Call rstFRS.Open("SELECT IDFRS FROM GRB_Fournisseur WHERE NomFournisseur = 'FOURNI PAR LE CLIENT'", g_connData, adOpenDynamic, adLockOptimistic)
      
185     iNoClient = rstFRS.Fields("IDFRS")

190     Call rstFRS.Close
195     Set rstFRS = Nothing

        'tant il y a des fournisseur de la piece, ajoute dans lister
200     Do While Not rstPieceFRS.EOF
205       If rstPieceFRS.Fields("IDFRS") = iNoClient Then
210         Call rstPieceFRS.MoveNext

215         If rstPieceFRS.EOF Then
220           Exit Do
225         End If
230       End If

          'on change la couleur de l'enregistrement selon la devise monétaire.
          'CAN = noir, USA ou ESP = bleu
235       If rstPieceFRS.Fields("DeviseMonétaire") = "CAN" Then
240         sDevise = "CAN"
245         lColor = COLOR_NOIR
250       Else
255         If rstPieceFRS.Fields("DeviseMonétaire") = "USA" Then
260           sDevise = "USA"
265           lColor = COLOR_BLEU
270         Else
275           sDevise = "SPA"
280           lColor = COLOR_BLEU
285         End If
290       End If
       
295       Set itmFRS = lvwfournisseur.ListItems.Add
       
300       itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = " "
305       itmFRS.SubItems(I_COL_FRS_ESCOMPTE) = " "
310       itmFRS.SubItems(I_COL_FRS_PRIX_NET) = " "
315       itmFRS.SubItems(I_COL_FRS_PRIX_SP) = " "
         
          'Nom du FRS
320       itmFRS.Text = rstPieceFRS.Fields("NomFournisseur")
           
325       itmFRS.Tag = rstPieceFRS.Fields("IDFRS")

330       itmFRS.ForeColor = lColor
      
          'Personne ressource
335       If Trim(rstPieceFRS.Fields("PERS_RESS")) <> vbNullString Then
340         Set rstContact = New ADODB.Recordset

345         Call rstContact.Open("SELECT NomContact FROM GRB_Contact WHERE IDContact = " & rstPieceFRS.Fields("PERS_RESS"), g_connData, adOpenDynamic, adLockOptimistic)
                
350         If Not rstContact.EOF Then
355           itmFRS.SubItems(I_COL_FRS_PERS_RESS) = rstContact.Fields("NomContact")

360           itmFRS.ListSubItems(I_COL_FRS_PERS_RESS).ForeColor = lColor
365         Else
370           itmFRS.SubItems(I_COL_FRS_PERS_RESS) = ""
375         End If
              
380         Call rstContact.Close
385         Set rstContact = Nothing
390       End If
                     
          'Date
395       If Not IsNull(rstPieceFRS.Fields("Date")) Then
400         itmFRS.SubItems(I_COL_FRS_DATE) = rstPieceFRS.Fields("Date")

405         itmFRS.ListSubItems(I_COL_FRS_DATE).ForeColor = lColor
410       Else
415         itmFRS.SubItems(I_COL_FRS_DATE) = vbNullString
420       End If
                          
          'Entrer par
425       If Not IsNull(rstPieceFRS.Fields("Entrer_Par")) Then
430         itmFRS.SubItems(I_COL_FRS_ENTRER_PAR) = rstPieceFRS.Fields("Entrer_Par")

435         itmFRS.ListSubItems(I_COL_FRS_ENTRER_PAR).ForeColor = lColor
440       Else
445         itmFRS.SubItems(I_COL_FRS_ENTRER_PAR) = vbNullString
450       End If
                                 
          'Valide
455       If Not IsNull(rstPieceFRS.Fields("Valide")) Then
460         itmFRS.SubItems(I_COL_FRS_VALIDE) = rstPieceFRS.Fields("Valide")

465         itmFRS.ListSubItems(I_COL_FRS_VALIDE).ForeColor = lColor
470       Else
475         itmFRS.SubItems(I_COL_FRS_VALIDE) = vbNullString
480       End If
                             
          'Prix listé
485       If rstPieceFRS.Fields("PRIX_LIST") <> vbNullString Then
490         itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = Conversion(Round(Replace(rstPieceFRS.Fields("PRIX_LIST"), ".", ","), 4), MODE_ARGENT, 4)

495         itmFRS.ListSubItems(I_COL_FRS_PRIX_LIST).ForeColor = lColor
500       End If
                              
          'Escompte
505       If rstPieceFRS.Fields("ESCOMPTE") <> vbNullString Then
510         itmFRS.SubItems(I_COL_FRS_ESCOMPTE) = Conversion(Replace(Replace(rstPieceFRS.Fields("ESCOMPTE"), "_", vbNullString), ".", ",") * 100, MODE_POURCENT)

515         itmFRS.ListSubItems(I_COL_FRS_ESCOMPTE).ForeColor = lColor
520       End If
      
          'Prix net
525       If rstPieceFRS.Fields("PRIX_NET") <> vbNullString Then
530         itmFRS.SubItems(I_COL_FRS_PRIX_NET) = Conversion(Round(Replace(rstPieceFRS.Fields("PRIX_NET"), ".", ","), 4), MODE_ARGENT, 4)

535         itmFRS.ListSubItems(I_COL_FRS_PRIX_NET).ForeColor = lColor
540       End If
    
          'Prix spécial
545       If rstPieceFRS.Fields("PRIX_SP") <> vbNullString Then
550         itmFRS.SubItems(I_COL_FRS_PRIX_SP) = Conversion(Round(rstPieceFRS.Fields("PRIX_SP"), 4), MODE_ARGENT, 4)

555         itmFRS.ListSubItems(I_COL_FRS_PRIX_SP).ForeColor = lColor
560       End If

          'Quoter
565       If rstPieceFRS.Fields("QUOTER") = True Then
570         itmFRS.SubItems(I_COL_FRS_QUOTER) = "Oui"
575       Else
580         itmFRS.SubItems(I_COL_FRS_QUOTER) = "Non"
585       End If

590       itmFRS.ListSubItems(I_COL_FRS_QUOTER).ForeColor = lColor
    
          'Pour garder en mémoire le prix d'origine, je le mets dans le
          'tag de la colonne Prix Listé
595       If itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = vbNullString Then
600         itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = " "
605       End If

610       If rstPieceFRS.Fields("PRIX_NET") <> vbNullString Then
615         If rstPieceFRS.Fields("PRIX_LIST") = "0,00" Or rstPieceFRS.Fields("PRIX_LIST") = "0" Then
620           itmFRS.ListSubItems(I_COL_FRS_PRIX_LIST).Tag = Replace(rstPieceFRS.Fields("PRIX_NET"), ".", ",")
625         Else
630           itmFRS.ListSubItems(I_COL_FRS_PRIX_LIST).Tag = Replace(rstPieceFRS.Fields("PRIX_LIST"), ".", ",")
635         End If
640       Else
645         itmFRS.ListSubItems(I_COL_FRS_PRIX_LIST).Tag = Replace(rstPieceFRS.Fields("PRIX_SP"), ".", ",")
650       End If

655       If itmFRS.SubItems(I_COL_FRS_PERS_RESS) = "" Then
660         itmFRS.SubItems(I_COL_FRS_PERS_RESS) = " "
665       End If

670       itmFRS.ListSubItems(I_COL_FRS_PERS_RESS).Tag = sDevise

675       Call rstPieceFRS.MoveNext
680     Loop
    
        'ferme la table
685     Call rstPieceFRS.Close
690     Set rstPieceFRS = Nothing

695     If m_bPieceInutile = False Then
700       Set itmFRS = lvwfournisseur.ListItems.Add

705       itmFRS.Text = "CHOISIR ULTÉRIEUREMENT"

710       itmFRS.SubItems(I_COL_FRS_PRIX_LIST) = " "
715       itmFRS.SubItems(I_COL_FRS_ESCOMPTE) = " "
720       itmFRS.SubItems(I_COL_FRS_PRIX_NET) = " "
725       itmFRS.SubItems(I_COL_FRS_PRIX_SP) = " "
730     End If

735     Exit Sub

AfficherErreur:

740     woups "frmAchat", "RemplirListViewFournisseur", Err, Erl
End Sub

Private Sub RemplirListViewInventaire()

5       On Error GoTo AfficherErreur

        'Rempli le listview des pièces à commander dans l'inventaire
10      Dim rstInv   As ADODB.Recordset
15      Dim itmInv   As ListItem
20      Dim lStock   As Long
25      Dim lMinimum As Long

        'Il faut vider le ListView avant de le remplir
30      Call lvwInventaire.ListItems.Clear
  
35      Set rstInv = New ADODB.Recordset
  
40      If m_eCatalogue = ELECTRIQUE Then
45        Call rstInv.Open("SELECT * FROM GRB_InventaireElec WHERE Minimum = True ORDER BY NoItem", g_connData, adOpenDynamic, adLockOptimistic)
50      Else
55        Call rstInv.Open("SELECT * FROM GRB_InventaireMec WHERE Minimum = True ORDER BY NoItem", g_connData, adOpenDynamic, adLockOptimistic)
60      End If
    
        'Tant que ce n'est pas la fin des enregistrements
65      Do While Not rstInv.EOF
70        If Not IsNull(rstInv.Fields("QuantitéStock")) Then
75          lStock = Replace(rstInv.Fields("QuantitéStock"), ".", ",")
80        Else
85          lStock = 0
90        End If
        
95        If Not IsNull(rstInv.Fields("QuantitéMinimum")) Then
100         lMinimum = rstInv.Fields("QuantitéMinimum")
105       Else
110         lMinimum = 0
115       End If
    
120       If lStock < lMinimum Then
            'On l'ajoute
125         Set itmInv = lvwInventaire.ListItems.Add
        
            'No piece
130         If Not IsNull(rstInv.Fields("NoItem")) Then
135           itmInv.Text = rstInv.Fields("NoItem")
140         Else
145           itmInv.Text = vbNullString
150         End If
        
            'Fabricant
155         If Not IsNull(rstInv.Fields("Manufacturier")) Then
160           itmInv.SubItems(I_COL_INV_MANUFACT) = rstInv.Fields("Manufacturier")
165         Else
170           itmInv.SubItems(I_COL_INV_MANUFACT) = vbNullString
175         End If
                   
            'Description
180         If Not IsNull(rstInv.Fields("Description")) Then
185           itmInv.SubItems(I_COL_INV_DESCR) = rstInv.Fields("Description")
190         Else
195           itmInv.SubItems(I_COL_INV_DESCR) = vbNullString
200         End If

            'Commentaire
205         If Not IsNull(rstInv.Fields("Commentaires")) Then
210           itmInv.SubItems(I_COL_INV_COMMENT) = rstInv.Fields("Commentaires")
215         Else
220           itmInv.SubItems(I_COL_INV_COMMENT) = ""
225         End If

            'Quantité en stock
230         itmInv.SubItems(I_COL_INV_QTE_STOCK) = lStock

            'Quantité minimum
235         itmInv.SubItems(I_COL_INV_QTE_MINIMUM) = lMinimum

            'Quantité à commander
240         If Not IsNull(rstInv.Fields("Commande")) Then
245           itmInv.SubItems(I_COL_INV_QTE_COMMANDE) = rstInv.Fields("Commande")
250         Else
255           itmInv.SubItems(I_COL_INV_QTE_COMMANDE) = vbNullString
260         End If
265       End If
      
270       Call rstInv.MoveNext
275     Loop
    
280     Call rstInv.Close
285     Set rstInv = Nothing

290     Exit Sub

AfficherErreur:

295     woups "frmAchat", "RemplirListViewInventaire", Err, Erl
End Sub


Private Sub RemplirListViewPieces()

5       On Error GoTo AfficherErreur

        'Rempli le listview des pièces selon la catégorie de pièce choisit
10      Dim rstPieces  As ADODB.Recordset
15      Dim itmPieces  As ListItem
20      Dim iIndex     As Integer
25      Dim bDebut     As Boolean
30      Dim sTri       As String
35      Dim sOrderBy   As String
40      Dim sCategorie As String
  
45      sTri = m_sTri
  
        'Il faut vider le ListView avant de le remplir
50      Call lvwPieces.ListItems.Clear

55      Set rstPieces = New ADODB.Recordset
  
60      Select Case cmbTri.ListIndex
          Case I_CMB_PIECE_GRB: sOrderBy = "PIECE_GRB"
65        Case I_CMB_PIECE:     sOrderBy = "PIECE"
70        Case I_CMB_FABRICANT: sOrderBy = "FABRICANT"
75        Case I_CMB_DESCR_FR:  sOrderBy = "DESC_FR"
80        Case I_CMB_DESCR_EN:  sOrderBy = "DESC_EN"
85      End Select
    
90      sCategorie = Replace(cmbCategorie.Text, "'", "''")
    
95      If m_eCatalogue = ELECTRIQUE Then
100       Call rstPieces.Open("SELECT * FROM GRB_CatalogueElec WHERE CATEGORIE = '" & sCategorie & "' ORDER BY " & sOrderBy, g_connData, adOpenDynamic, adLockOptimistic)
105     Else
110       Call rstPieces.Open("SELECT * FROM GRB_CatalogueMec WHERE CATEGORIE = '" & sCategorie & "' ORDER BY " & sOrderBy, g_connData, adOpenDynamic, adLockOptimistic)
115     End If
    
120     iIndex = 1
    
        'Tant que ce n'est pas la fin des enregistrements
125     Do While Not rstPieces.EOF
130       If rstPieces.Fields("PIECE") <> vbNullString And rstPieces.Fields("FABRICANT") <> vbNullString Then
            'Si il y a une recherche à faire
135         If sTri <> vbNullString Then
140           bDebut = False
      
              'Selon la colonne
145           Select Case m_iCol
                'Si c'est la colonne PIECE_GRB
                Case I_COL_PIECES_PIECE_GRB:
                  'Si la PIECE_GRB contient la recherche
150               If InStr(1, UCase(rstPieces.Fields("PIECE_GRB")), UCase(sTri)) > 0 Then
                    'On met la variable à true pour l'ajouter au début
155                 bDebut = True
160               End If
                      
                'Si c'est la colonne No. d'item
165             Case I_COL_PIECES_NO_ITEM:
                  'Si le no. d'item contient la recherche
170               If InStr(1, UCase(rstPieces.Fields("PIECE")), UCase(sTri)) > 0 Then
                    'On met la variable à true pour l'ajouter au début
175                 bDebut = True
180               End If
        
                'Si c'est la colonne Manufacturier
185             Case I_COL_PIECES_MANUFACT:
                  'Si le manufacturier contient la recherche
190               If InStr(1, UCase(rstPieces.Fields("FABRICANT")), UCase(sTri)) > 0 Then
                    'On met la variable à true pour l'ajouter au début
195                 bDebut = True
200               End If
            
                'Si c'est la colonne No. d'item
205             Case I_COL_PIECES_DESCR_FR:
                  'Si la description française contient la recherche
210               If InStr(1, UCase(rstPieces.Fields("DESC_FR")), UCase(sTri)) > 0 Then
                    'On met la variable à true pour l'ajouter au début
215                 bDebut = True
220               End If
            
                'Si c'est la colonne No. d'item
225             Case I_COL_PIECES_DESCR_EN:
                  'Si la description anglaise contient la recherche
230               If InStr(1, UCase(rstPieces.Fields("DESC_EN")), UCase(sTri)) > 0 Then
                    'On met la variable à true pour l'ajouter au début
235                 bDebut = True
240               End If
245           End Select
      
250           If bDebut = True Then
255             Set itmPieces = lvwPieces.ListItems.Add(iIndex)
          
260             iIndex = iIndex + 1
265           Else
270             Set itmPieces = lvwPieces.ListItems.Add
275           End If
280         Else
285           Set itmPieces = lvwPieces.ListItems.Add
290         End If
        
            'Piece_GRB
295         If Not IsNull(rstPieces.Fields("PIECE_GRB")) Then
300           itmPieces.Text = rstPieces.Fields("PIECE_GRB")
305         Else
310           itmPieces.Text = vbNullString
315         End If
        
            'No piece
320         If Not IsNull(rstPieces.Fields("PIECE")) Then
325           itmPieces.SubItems(I_COL_PIECES_NO_ITEM) = rstPieces.Fields("PIECE")
330         Else
335           itmPieces.SubItems(I_COL_PIECES_NO_ITEM) = vbNullString
340         End If
        
            'Fabricant
345         If Not IsNull(rstPieces.Fields("FABRICANT")) Then
350           itmPieces.SubItems(I_COL_PIECES_MANUFACT) = rstPieces.Fields("FABRICANT")
355         Else
360           itmPieces.SubItems(I_COL_PIECES_MANUFACT) = vbNullString
365         End If
                   
            'Description en francais
370         If Not IsNull(rstPieces.Fields("DESC_FR")) Then
375           itmPieces.SubItems(I_COL_PIECES_DESCR_FR) = rstPieces.Fields("DESC_FR")
380         Else
385           itmPieces.SubItems(I_COL_PIECES_DESCR_FR) = vbNullString
390         End If
        
            'Description en anglais
395         If Not IsNull(rstPieces.Fields("DESC_EN")) Then
400           itmPieces.SubItems(I_COL_PIECES_DESCR_EN) = rstPieces.Fields("DESC_EN")
405         Else
410           itmPieces.SubItems(I_COL_PIECES_DESCR_EN) = vbNullString
415         End If
420       End If

          'Commentaire
425       If Not IsNull(rstPieces.Fields("COMMENTAIRE")) Then
430         itmPieces.SubItems(I_COL_PIECES_COMMENT) = rstPieces.Fields("COMMENTAIRE")
435       Else
440         itmPieces.SubItems(I_COL_PIECES_COMMENT) = ""
445       End If
      
450       Call rstPieces.MoveNext
455     Loop
    
460     Call rstPieces.Close
465     Set rstPieces = Nothing

470     Exit Sub

AfficherErreur:

475     woups "frmAchat", "RemplirListViewPieces", Err, Erl
End Sub

Private Sub RemplirListViewAchat()

5       On Error GoTo AfficherErreur

        'Remplis les pièces de l'achat avec la BD
10      Dim rstAchat    As ADODB.Recordset
15      Dim rstFRS      As ADODB.Recordset
20      Dim itmAchat    As ListItem
25      Dim sIDAchat    As String
30      Dim iIndexAchat As Integer
35      Dim lColor      As Long
40      Dim bBold       As Boolean
    
45      Call lvwAchat.ListItems.Clear
  
50      sIDAchat = Left$(txtNoAchat.Text, 9)
  
55      iIndexAchat = CInt(Right$(txtNoAchat.Text, 3))
  
60      Set rstAchat = New ADODB.Recordset
  
65      Call rstAchat.Open("SELECT * FROM GRB_Achat_Pieces WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat & " ORDER BY NuméroLigne", g_connData, adOpenDynamic, adLockOptimistic)
  
70      Do While Not rstAchat.EOF
75        bBold = False

80        If rstAchat.Fields("Retour") = True Then
85          lColor = COLOR_ROUGE
90        Else
95          If rstAchat.Fields("Recu") = True Then
100           lColor = COLOR_GRIS 'Gris
105         Else
110           If rstAchat.Fields("Inutile") = True Then
115             lColor = COLOR_BRUN
120           Else
125             If rstAchat.Fields("IDFRS") = 0 Then
130               lColor = COLOR_MAGENTA
135             Else
140               If rstAchat.Fields("Commandé") = True Then
145                 lColor = COLOR_ORANGE     'COLOR_ORANGE
150               Else
155                 If rstAchat.Fields("CommandeAnnulée") = True Then
160                   lColor = COLOR_VERT_FORET
165                   bBold = True
170                 Else
175                   lColor = COLOR_NOIR
180                 End If
185               End If
190             End If
195           End If
200         End If
205       End If

210       Set itmAchat = lvwAchat.ListItems.Add
          
          'Quantité
215       If Not IsNull(rstAchat.Fields("Qté")) Then
220         itmAchat.Text = rstAchat.Fields("Qté")
225       Else
230         itmAchat.Text = vbNullString
235       End If

240       itmAchat.ForeColor = lColor
245       itmAchat.Bold = bBold
   
250       itmAchat.Tag = rstAchat.Fields("DateRéception")
    
          'Numéro d'item
255       If Not IsNull(rstAchat.Fields("PIECE")) Then
260         itmAchat.SubItems(I_COL_ACHAT_PIECE) = rstAchat.Fields("PIECE")
265       Else
270         itmAchat.SubItems(I_COL_ACHAT_PIECE) = vbNullString
275       End If

280       itmAchat.ListSubItems(I_COL_ACHAT_PIECE).ForeColor = lColor
285       itmAchat.ListSubItems(I_COL_ACHAT_PIECE).Bold = bBold
            
          'Description en francais
290       If Not IsNull(rstAchat.Fields("DESC_FR")) Then
295         itmAchat.SubItems(I_COL_ACHAT_DESCR) = rstAchat.Fields("DESC_FR")
300       Else
305         itmAchat.SubItems(I_COL_ACHAT_DESCR) = vbNullString
310       End If

315       itmAchat.ListSubItems(I_COL_ACHAT_DESCR).ForeColor = lColor
320       itmAchat.ListSubItems(I_COL_ACHAT_DESCR).Bold = bBold
    
          'On met la description en anglais dans le tag de la description en francais
325       If Not IsNull(rstAchat.Fields("Desc_EN")) Then
330         itmAchat.ListSubItems(I_COL_ACHAT_DESCR).Tag = rstAchat.Fields("Desc_EN")
335       Else
340         itmAchat.ListSubItems(I_COL_ACHAT_DESCR).Tag = vbNullString
345       End If
   
          'Fabricant
350       If Not IsNull(rstAchat.Fields("Manufact")) Then
355         itmAchat.SubItems(I_COL_ACHAT_MANUFACT) = rstAchat.Fields("Manufact")
360       Else
365         itmAchat.SubItems(I_COL_ACHAT_MANUFACT) = vbNullString
370       End If

375       itmAchat.ListSubItems(I_COL_ACHAT_MANUFACT).ForeColor = lColor
380       itmAchat.ListSubItems(I_COL_ACHAT_MANUFACT).Bold = bBold

385       itmAchat.ListSubItems(I_COL_ACHAT_MANUFACT).Tag = rstAchat.Fields("NoRetour")
    
          'Prix listé
390       If Trim(rstAchat.Fields("Prix_List")) <> vbNullString Then
395         itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = Conversion(rstAchat.Fields("Prix_list"), MODE_ARGENT, 4)
400       Else
405         itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = " "
410       End If

415       itmAchat.ListSubItems(I_COL_ACHAT_PRIX_LIST).ForeColor = lColor

420       itmAchat.ListSubItems(I_COL_ACHAT_PRIX_LIST).Bold = bBold

425       itmAchat.ListSubItems(I_COL_ACHAT_PRIX_LIST).Tag = rstAchat.Fields("PrixOrigine")
      
          'Escompte
430       If Trim(rstAchat.Fields("Escompte")) <> vbNullString Then
435         itmAchat.SubItems(I_COL_ACHAT_ESCOMPTE) = Conversion(rstAchat.Fields("Escompte"), MODE_POURCENT)
440       Else
445         itmAchat.SubItems(I_COL_ACHAT_ESCOMPTE) = " "
450       End If

455       itmAchat.ListSubItems(I_COL_ACHAT_ESCOMPTE).ForeColor = lColor
460       itmAchat.ListSubItems(I_COL_ACHAT_ESCOMPTE).Bold = bBold
    
          'Prix net
465       If Trim(rstAchat.Fields("Prix_net")) <> vbNullString Then
470         itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(rstAchat.Fields("Prix_net"), MODE_ARGENT, 4)
475       Else
480         itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = " "
485       End If

490       itmAchat.ListSubItems(I_COL_ACHAT_PRIX_NET).ForeColor = lColor
495       itmAchat.ListSubItems(I_COL_ACHAT_PRIX_NET).Bold = bBold
         
          'Fournisseur
500       If Not IsNull(rstAchat.Fields("IDFRS")) Then
505         If rstAchat.Fields("IDFRS") <> 0 Then
510           Set rstFRS = New ADODB.Recordset

515           Call rstFRS.Open("SELECT NomFournisseur FROM GRB_Fournisseur WHERE IDFRS = " & rstAchat.Fields("IDFRS"), g_connData, adOpenDynamic, adLockOptimistic)
   
              'On affiche le nom dans la colonne
520           If Not rstFRS.EOF Then
525             itmAchat.SubItems(I_COL_ACHAT_DISTRIB) = rstFRS.Fields("NomFournisseur")
530           Else
535             itmAchat.SubItems(I_COL_ACHAT_DISTRIB) = ""
540           End If
        
              'On affiche l'Id dans le tag
545           itmAchat.ListSubItems(I_COL_ACHAT_DISTRIB).Tag = rstAchat.Fields("IDFRS")
        
550           Call rstFRS.Close
555           Set rstFRS = Nothing
560         Else
565           itmAchat.SubItems(I_COL_ACHAT_DISTRIB) = " "
570         End If
575       Else
580         itmAchat.SubItems(I_COL_ACHAT_DISTRIB) = vbNullString
585       End If

590       itmAchat.ListSubItems(I_COL_ACHAT_DISTRIB).ForeColor = lColor
595       itmAchat.ListSubItems(I_COL_ACHAT_DISTRIB).Bold = bBold
    
          'Prix total
600       If rstAchat.Fields("Prix_total") <> vbNullString Then
605         itmAchat.SubItems(I_COL_ACHAT_TOTAL) = Conversion(Round(rstAchat.Fields("Prix_total"), 2), MODE_ARGENT)
610       Else
615         itmAchat.SubItems(I_COL_ACHAT_TOTAL) = " "
620       End If

625       itmAchat.ListSubItems(I_COL_ACHAT_TOTAL).ForeColor = lColor
630       itmAchat.ListSubItems(I_COL_ACHAT_TOTAL).Bold = bBold

635       itmAchat.ListSubItems(I_COL_ACHAT_TOTAL).Tag = rstAchat.Fields("Devise")

          'Date Commande
640       If rstAchat.Fields("DateCommande") <> vbNullString Then
645         itmAchat.SubItems(I_COL_ACHAT_DATE_COMMANDE) = rstAchat.Fields("DateCommande")
650       Else
655         itmAchat.SubItems(I_COL_ACHAT_DATE_COMMANDE) = ""
660       End If

665       itmAchat.ListSubItems(I_COL_ACHAT_DATE_COMMANDE).ForeColor = lColor
670       itmAchat.ListSubItems(I_COL_ACHAT_DATE_COMMANDE).Bold = bBold

          'Date Requise
675       If rstAchat.Fields("DateRequise") <> vbNullString Then
680         itmAchat.SubItems(I_COL_ACHAT_DATE_REQUISE) = rstAchat.Fields("DateRequise")
685       Else
690         itmAchat.SubItems(I_COL_ACHAT_DATE_REQUISE) = ""
695       End If

700       itmAchat.ListSubItems(I_COL_ACHAT_DATE_REQUISE).ForeColor = lColor
705       itmAchat.ListSubItems(I_COL_ACHAT_DATE_REQUISE).Bold = bBold

710       Call rstAchat.MoveNext
715     Loop
  
720     Call rstAchat.Close
725     Set rstAchat = Nothing

730     Exit Sub

AfficherErreur:

735     woups "frmAchat", "RemplirListViewAchat", Err, Erl
End Sub

Private Sub lvwFournisseur_DblClick()

5       On Error GoTo AfficherErreur

10      If m_bPieceInutile = True Then
15        Call ChoisirFournisseurMateriel
20      Else
25        Call ChoisirFournisseur
30      End If

35      Exit Sub

AfficherErreur:

40      woups "frmAchat", "lvwFournisseur_DblClick", Err, Erl
End Sub

Private Sub ChoisirFournisseur()

5       On Error GoTo AfficherErreur

        'On ajoute la pièce dans lvwAchat
10      Dim rstConfig As ADODB.Recordset
15      Dim sTauxUSA  As String
20      Dim sTauxSPA  As String
25      Dim sQuantite As String
30      Dim itmAchat  As ListItem
35      Dim lColor    As Long
     
        'Saisie de la quantité
40      sQuantite = InputBox("Quelle est la quantité?", , m_sQuantite)

45      sQuantite = Replace(sQuantite, ".", ",")

50      Set rstConfig = New ADODB.Recordset

55      Call rstConfig.Open("SELECT TauxAmericain, TauxEspagnol FROM GRB_Config", g_connData, adOpenForwardOnly, adLockReadOnly)

60      sTauxUSA = rstConfig.Fields("TauxAmericain")
65      sTauxSPA = rstConfig.Fields("TauxEspagnol")

70      Call rstConfig.Close
75      Set rstConfig = Nothing
    
80      If sQuantite <> vbNullString Then
85        If Not IsNumeric(sQuantite) Then
90          Call MsgBox("Quantité non numérique!", vbOKOnly, "Erreur")
      
95          Exit Sub
100       End If
105     Else
110       Exit Sub
115     End If
    
120     Set itmAchat = lvwAchat.ListItems.Add
          
125     If lvwfournisseur.SelectedItem.Text = "CHOISIR ULTÉRIEUREMENT" Then
130       lColor = COLOR_MAGENTA
135     Else
140       lColor = COLOR_NOIR
145     End If
                 
        'Quantité
150     itmAchat.Text = sQuantite
155     itmAchat.ForeColor = lColor
  
        'Numéro d'item
160     If m_bRecherchePiece = True Then
165       itmAchat.SubItems(I_COL_ACHAT_PIECE) = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM)
170       itmAchat.SubItems(I_COL_ACHAT_DESCR) = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_DESCR_FR)
175       itmAchat.SubItems(I_COL_ACHAT_MANUFACT) = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_MANUFACT)
180       itmAchat.ListSubItems(I_COL_ACHAT_DESCR).Tag = lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_DESCR_EN)
185     Else
190       If m_bInventaire = True Then
195         itmAchat.SubItems(I_COL_ACHAT_PIECE) = lvwInventaire.SelectedItem.Text
200         itmAchat.SubItems(I_COL_ACHAT_DESCR) = lvwInventaire.SelectedItem.SubItems(I_COL_INV_DESCR)
205         itmAchat.SubItems(I_COL_ACHAT_MANUFACT) = lvwInventaire.SelectedItem.SubItems(I_COL_INV_MANUFACT)
210       Else
215         itmAchat.SubItems(I_COL_ACHAT_PIECE) = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_NO_ITEM)
220         itmAchat.SubItems(I_COL_ACHAT_DESCR) = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_DESCR_FR)
225         itmAchat.SubItems(I_COL_ACHAT_MANUFACT) = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_MANUFACT)
230         itmAchat.ListSubItems(I_COL_ACHAT_DESCR).Tag = lvwPieces.SelectedItem.SubItems(I_COL_PIECES_DESCR_EN)
235       End If
240     End If

245     itmAchat.ListSubItems(I_COL_ACHAT_PIECE).ForeColor = lColor
250     itmAchat.ListSubItems(I_COL_ACHAT_DESCR).ForeColor = lColor
255     itmAchat.ListSubItems(I_COL_ACHAT_MANUFACT).ForeColor = lColor
     
        'Prix listé
260     If Trim$(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_LIST)) <> vbNullString Then
265       If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "USA" Then
270         itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_LIST)) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
275       Else
280         If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "SPA" Then
285           itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_LIST)) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
290         Else
295           itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = Conversion(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_LIST), MODE_ARGENT, 4)
300         End If
305       End If
310     Else
315       itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = Conversion("0", MODE_ARGENT, 4)
320     End If

325     itmAchat.ListSubItems(I_COL_ACHAT_PRIX_LIST).Tag = lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PRIX_LIST).Tag

330     itmAchat.ListSubItems(I_COL_ACHAT_PRIX_LIST).ForeColor = lColor
       
335     If Trim$(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_NET)) <> vbNullString Then
340       If Trim$(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_ESCOMPTE)) <> vbNullString Then
345         itmAchat.SubItems(I_COL_ACHAT_ESCOMPTE) = Conversion(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_ESCOMPTE), MODE_POURCENT)
350       Else
355         itmAchat.SubItems(I_COL_ACHAT_ESCOMPTE) = Conversion("0", MODE_POURCENT)
360       End If

365       If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "USA" Then
370         itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_NET)) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
375       Else
380         If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "SPA" Then
385           itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_NET)) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
390         Else
395           itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_NET), MODE_ARGENT, 4)
400         End If
405       End If
410     Else
415       If Trim$(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_SP)) <> vbNullString Then
420         If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "USA" Then
425           itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_SP)) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
430         Else
435           If lvwfournisseur.SelectedItem.ListSubItems(I_COL_FRS_PERS_RESS).Tag = "SPA" Then
440             itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(CStr(Round(CDbl(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_SP)) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
445           Else
450             itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(lvwfournisseur.SelectedItem.SubItems(I_COL_FRS_PRIX_SP), MODE_ARGENT, 4)
455           End If
460         End If
465       Else
470         itmAchat.SubItems(I_COL_ACHAT_ESCOMPTE) = Conversion("0", MODE_POURCENT)
475         itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion("0", MODE_ARGENT, 4)
480       End If
485     End If

490     itmAchat.ListSubItems(I_COL_ACHAT_ESCOMPTE).ForeColor = lColor
495     itmAchat.ListSubItems(I_COL_ACHAT_PRIX_NET).ForeColor = lColor

500     If lvwfournisseur.SelectedItem.Text = "CHOISIR ULTÉRIEUREMENT" Then
505       itmAchat.SubItems(I_COL_ACHAT_DISTRIB) = " "
510       itmAchat.ListSubItems(I_COL_ACHAT_DISTRIB).Tag = 0
515     Else
520       itmAchat.SubItems(I_COL_ACHAT_DISTRIB) = lvwfournisseur.SelectedItem.Text
525       itmAchat.ListSubItems(I_COL_ACHAT_DISTRIB).Tag = lvwfournisseur.SelectedItem.Tag
530     End If

535     itmAchat.ListSubItems(I_COL_ACHAT_DISTRIB).ForeColor = lColor
  
        'Prix total
540     itmAchat.SubItems(I_COL_ACHAT_TOTAL) = Conversion(Round(itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) * itmAchat.Text, 2), MODE_ARGENT)
545     itmAchat.ListSubItems(I_COL_ACHAT_TOTAL).ForeColor = lColor
      
        'Calcul des prix
550     Call CalculerPrix
  
        'On cache le listview
555     frafournisseur.Visible = False

560     Exit Sub

AfficherErreur:

565     woups "frmAchat", "ChoisirFournisseur", Err, Erl
End Sub

Private Sub ChoisirFournisseurMateriel()

5       On Error GoTo AfficherErreur

        'On ajoute la pièce en négatif dans le listview
10      Dim sQuantite  As String
15      Dim itmAncien  As ListItem
20      Dim itmNouveau As ListItem
  
        'Saisie de la quantité
25      sQuantite = InputBox("Quelle est la quantité?")

30      sQuantite = Replace(sQuantite, ".", ",")
    
35      If sQuantite <> vbNullString Then
40        If Not IsNumeric(sQuantite) Then
45          Call MsgBox("Quantité non numérique!", vbOKOnly, "Erreur")
      
50          Exit Sub
55        End If
60      Else
65        Exit Sub
70      End If

75      If CDbl(sQuantite) <= CDbl(lvwAchat.SelectedItem.Text) Then
80        Set itmAncien = lvwAchat.SelectedItem
85        Set itmNouveau = lvwAchat.ListItems.Add(itmAncien.Index + 1)
  
90        itmNouveau.Checked = itmAncien.Checked
  
95        itmNouveau.Text = "-" & sQuantite
                                                                                                         
          'No d'item
100       itmNouveau.SubItems(I_COL_ACHAT_PIECE) = itmAncien.SubItems(I_COL_ACHAT_PIECE)
 
          'On met la description en francais dans la colonne et la description en anglais
          'dans le tag
105       itmNouveau.SubItems(I_COL_ACHAT_DESCR) = itmAncien.SubItems(I_COL_ACHAT_DESCR)
110       itmNouveau.ListSubItems(I_COL_ACHAT_DESCR).Tag = itmAncien.ListSubItems(I_COL_ACHAT_DESCR).Tag
          
          'On met le nom du fabricant dans la colonne et l'ordre de la section dans le tag
115       itmNouveau.SubItems(I_COL_ACHAT_MANUFACT) = itmAncien.SubItems(I_COL_ACHAT_MANUFACT)

          'Prix listé
120       itmNouveau.SubItems(I_COL_ACHAT_PRIX_LIST) = itmAncien.SubItems(I_COL_ACHAT_PRIX_LIST)

125       itmNouveau.ListSubItems(I_COL_ACHAT_PRIX_LIST).Tag = itmAncien.ListSubItems(I_COL_ACHAT_PRIX_LIST).Tag

130       itmNouveau.SubItems(I_COL_ACHAT_ESCOMPTE) = itmAncien.SubItems(I_COL_ACHAT_ESCOMPTE)

135       itmNouveau.SubItems(I_COL_ACHAT_PRIX_NET) = itmAncien.SubItems(I_COL_ACHAT_PRIX_NET)
            
          'On met le fournisseur dans la colonne et l'id dans le tag
140       itmNouveau.SubItems(I_COL_ACHAT_DISTRIB) = lvwfournisseur.SelectedItem.Text
145       itmNouveau.ListSubItems(I_COL_ACHAT_DISTRIB).Tag = lvwfournisseur.SelectedItem.Tag
      
150       itmNouveau.SubItems(I_COL_ACHAT_TOTAL) = Conversion(Round(CDbl(itmNouveau.Text) * CDbl(itmNouveau.SubItems(I_COL_ACHAT_PRIX_NET)), 2), MODE_ARGENT)

155       itmNouveau.SubItems(I_COL_ACHAT_DATE_COMMANDE) = " "
160       itmNouveau.SubItems(I_COL_ACHAT_DATE_REQUISE) = " "

165       itmNouveau.ForeColor = COLOR_BRUN
170       itmNouveau.ListSubItems(I_COL_ACHAT_DATE_COMMANDE).ForeColor = COLOR_BRUN
175       itmNouveau.ListSubItems(I_COL_ACHAT_DATE_REQUISE).ForeColor = COLOR_BRUN
180       itmNouveau.ListSubItems(I_COL_ACHAT_DESCR).ForeColor = COLOR_BRUN
185       itmNouveau.ListSubItems(I_COL_ACHAT_DISTRIB).ForeColor = COLOR_BRUN
190       itmNouveau.ListSubItems(I_COL_ACHAT_ESCOMPTE).ForeColor = COLOR_BRUN
195       itmNouveau.ListSubItems(I_COL_ACHAT_MANUFACT).ForeColor = COLOR_BRUN
200       itmNouveau.ListSubItems(I_COL_ACHAT_PIECE).ForeColor = COLOR_BRUN
205       itmNouveau.ListSubItems(I_COL_ACHAT_PRIX_LIST).ForeColor = COLOR_BRUN
210       itmNouveau.ListSubItems(I_COL_ACHAT_PRIX_NET).ForeColor = COLOR_BRUN
215       itmNouveau.ListSubItems(I_COL_ACHAT_TOTAL).ForeColor = COLOR_BRUN

          'Calcul des prix
220       Call CalculerPrix
  
          'On cache le listview
225       frafournisseur.Visible = False

230       m_bPieceInutile = False
  
          'Resélectionne le premier élément du listview
235       If lvwAchat.ListItems.count > 0 Then
240         lvwAchat.ListItems(1).Selected = True
245       End If
250     Else
255       Call MsgBox("Quantité trop grande!", vbOKOnly, "Erreur")
260     End If

265     Exit Sub

AfficherErreur:

270     woups "frmAchat", "ChoisirFournisseurMateriel", Err, Erl
End Sub

Private Sub CalculerPrix()

5       On Error GoTo AfficherErreur

10      Dim dblTotal  As Double
15      Dim iCompteur As Integer
  
20      If lvwAchat.ListItems.count > 0 Then
25        For iCompteur = 1 To lvwAchat.ListItems.count
30          dblTotal = dblTotal + CDbl(Conversion(lvwAchat.ListItems(iCompteur).SubItems(I_COL_ACHAT_TOTAL), MODE_PAS_FORMAT))
35        Next
    
40        txtPrixTotal.Text = Conversion(dblTotal, MODE_ARGENT)
45      Else
50        txtPrixTotal.Text = Conversion(0, MODE_ARGENT)
55      End If

60      Exit Sub

AfficherErreur:

65      woups "frmAchat", "CalculerPrix", Err, Erl
End Sub

Private Sub lvwFournisseur_LostFocus()

5       On Error GoTo AfficherErreur

        'On cache le Frame contenant le ListView si le ListView perd le focus
10      frafournisseur.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmAchat", "lvwFournisseur_LostFocus", Err, Erl
End Sub

Private Sub lvwPieces_DblClick()

5       On Error GoTo AfficherErreur

10      m_bInventaire = False
15      m_bRecherchePiece = -False

        'On affiche lvwFournisseur selon la pièce choisie
        'Rempli les fournisseurs de la pièce choisie
20      Call AfficherListeFournisseurs
  
        'Si le listview n'est pas vide
25      If lvwfournisseur.ListItems.count = 0 Then
30        If MsgBox("Il n'y a aucun fournisseur pour cette pièce!" & vbNewLine & "Voulez-vous en ajouter?", vbYesNo, "Erreur") = vbYes Then
35          Screen.MousePointer = vbHourglass
      
            'On ouvre le catalogue sur cet enregistrement
40          If m_eCatalogue = ELECTRIQUE Then
45            Call FrmCatalogueElec.AfficherForm(cmbCategorie.Text, lvwPieces.SelectedItem.SubItems(I_COL_PIECES_MANUFACT), lvwPieces.SelectedItem.SubItems(I_COL_PIECES_NO_ITEM))
50          Else
55            Call FrmCatalogueMec.AfficherForm(cmbCategorie.Text, lvwPieces.SelectedItem.SubItems(I_COL_PIECES_MANUFACT), lvwPieces.SelectedItem.SubItems(I_COL_PIECES_NO_ITEM))
60          End If
      
65          Screen.MousePointer = vbDefault
      
            'On rappelle la méthode
70          Call AfficherListeFournisseurs
75        End If
80      End If

85      Exit Sub

AfficherErreur:

90      woups "frmAchat", "lvwPieces_DblClick", Err, Erl
End Sub

Private Sub lvwInventaire_DblClick()

5       On Error GoTo AfficherErreur

        'On affiche lvwFournisseur selon la pièce choisie
        'Rempli les fournisseurs de la pièce choisie
        
10      m_bInventaire = True
15      m_bRecherchePiece = False
        
20      Call AfficherListeFournisseurs
  
        'Si le listview n'est pas vide
25      If lvwfournisseur.ListItems.count = 1 Then
30        If MsgBox("Il n'y a aucun fournisseur pour cette pièce!" & vbNewLine & "Voulez-vous en ajouter?", vbYesNo, "Erreur") = vbYes Then
35          Screen.MousePointer = vbHourglass
      
            'On ouvre le catalogue sur cet enregistrement
40          If m_eCatalogue = ELECTRIQUE Then
45            Call FrmCatalogueElec.AfficherForm(cmbCategorie.Text, lvwInventaire.SelectedItem.SubItems(I_COL_INV_MANUFACT), lvwInventaire.SelectedItem.SubItems(I_COL_INV_NO_ITEM))
50          Else
55            Call FrmCatalogueMec.AfficherForm(cmbCategorie.Text, lvwInventaire.SelectedItem.SubItems(I_COL_INV_MANUFACT), lvwInventaire.SelectedItem.SubItems(I_COL_INV_NO_ITEM))
60          End If
      
65          Screen.MousePointer = vbDefault
      
            'On rappelle la méthode
70          Call AfficherListeFournisseurs
75        End If
80      End If

85      fraInventaire.Visible = False

90      Exit Sub

AfficherErreur:

95      woups "frmAchat", "lvwInventaire_DblClick", Err, Erl
End Sub

Private Sub AfficherListeFournisseurs()

5       On Error GoTo AfficherErreur

        'Méthode qui sert à afficher la liste des fournisseurs
        'Affiche le frame seulement s'il y a des items dans le ListView
10      Call RemplirListViewFournisseur
  
15      If m_bInventaire = True Then
20        m_sQuantite = lvwInventaire.SelectedItem.SubItems(I_COL_INV_QTE_COMMANDE)
25      Else
30        m_sQuantite = vbNullString
35      End If
  
40      If lvwfournisseur.ListItems.count > 1 Then
45        If m_bRecherchePiece = True Then
50          fraPieceTrouve.Visible = False
55        End If

60        frafournisseur.Visible = True
65        Call lvwfournisseur.SetFocus
70      End If

75      Exit Sub

AfficherErreur:

80      woups "frmAchat", "AfficherListeFournisseurs", Err, Erl
End Sub

Private Sub lvwAchat_KeyDown(KeyCode As Integer, Shift As Integer)

5       On Error GoTo AfficherErreur

        'S'il est en mode ajout ou modif
10      If m_eMode = MODE_AJOUT_MODIF Then
          'Si le listView n'est pas vide
15        If lvwAchat.ListItems.count > 0 Then
            'Si la touche pesée est Delete
20          If KeyCode = vbKeyDelete Then
              'On l'efface
25            Call lvwAchat.ListItems.Remove(lvwAchat.SelectedItem.Index)
        
30            Call CalculerPrix
35          End If
40        End If
45      End If

50      Exit Sub

AfficherErreur:

55      woups "frmAchat", "lvwAchat_KeyDown", Err, Erl
End Sub

Private Sub cmdAnnulerPrix_Click()

5       On Error GoTo AfficherErreur

10      fraPrixPiece.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmAchat", "cmdAnnulerPrix_Click", Err, Erl
End Sub

Private Sub cmdOKPrix_Click()
        'Écrit les prix dans le listview
5       On Error GoTo AfficherErreur

10      Dim rstConfig As ADODB.Recordset
15      Dim itmAchat  As ListItem
20      Dim itmAvant  As ListItem
25      Dim bPrixSpecial As Boolean
30      Dim lColor    As Long
35      Dim iCompteur As Integer
40      Dim sQuantite As String
45      Dim sPiece    As String
50      Dim sTauxUSA  As String
55      Dim sTauxSPA  As String

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

245       Set itmAvant = lvwAchat.ListItems(CInt(fraPrixPiece.Tag))
250       Set itmAchat = lvwAchat.ListItems.Add(CInt(fraPrixPiece.Tag) + 1)

255       lColor = itmAvant.ListSubItems(I_COL_ACHAT_PIECE).ForeColor

260       itmAchat.Checked = itmAvant.Checked

          'Quantité
265       itmAchat.Text = "-" & itmAvant.Text

          'No d'item
270       itmAchat.SubItems(I_COL_ACHAT_PIECE) = itmAvant.SubItems(I_COL_ACHAT_PIECE)

          'On met la description en francais dans la colonne et la description en anglais
          'dans le tag
275       itmAchat.SubItems(I_COL_ACHAT_DESCR) = itmAvant.SubItems(I_COL_ACHAT_DESCR)
280       itmAchat.ListSubItems(I_COL_ACHAT_DESCR).Tag = itmAvant.ListSubItems(I_COL_ACHAT_DESCR).Tag

          'On met le nom du fabricant dans la col-nne et l'ordre de la section dans le tag
285       itmAchat.SubItems(I_COL_ACHAT_MANUFACT) = itmAvant.SubItems(I_COL_ACHAT_MANUFACT)

          'Prix listé
290       itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = itmAvant.SubItems(I_COL_ACHAT_PRIX_LIST)

295       itmAchat.ListSubItems(I_COL_ACHAT_PRIX_LIST).Tag = itmAvant.ListSubItems(I_COL_ACHAT_PRIX_LIST).Tag

300       itmAchat.SubItems(I_COL_ACHAT_ESCOMPTE) = itmAvant.SubItems(I_COL_ACHAT_ESCOMPTE)

305       itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = itmAvant.SubItems(I_COL_ACHAT_PRIX_NET)

          'On met le fournisseur dans la colonne et l'id dans le tag
310       itmAchat.SubItems(I_COL_ACHAT_DISTRIB) = itmAvant.SubItems(I_COL_ACHAT_DISTRIB)
315       itmAchat.ListSubItems(I_COL_ACHAT_DISTRIB).Tag = itmAvant.ListSubItems(I_COL_ACHAT_DISTRIB).Tag

          'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
320       itmAchat.SubItems(I_COL_ACHAT_TOTAL) = "-" & itmAvant.SubItems(I_COL_ACHAT_TOTAL)

          'Ajout de l'enregistrement avec le nouveau prix
325       Set itmAchat = lvwAchat.ListItems.Add(CInt(fraPrixPiece.Tag) + 2)

330       itmAchat.Checked = itmAvant.Checked

          'Quantité
335       itmAchat.Text = sQuantite

340       itmAchat.ForeColor = lColor
          
          'No d'item
345       itmAchat.SubItems(I_COL_ACHAT_PIECE) = itmAvant.SubItems(I_COL_ACHAT_PIECE)

350       itmAchat.ListSubItems(I_COL_ACHAT_PIECE).ForeColor = lColor

          'On met la description en francais dans la colonne et la description en anglais
          'dans le tag
355       itmAchat.SubItems(I_COL_ACHAT_DESCR) = itmAvant.SubItems(I_COL_ACHAT_DESCR)
360       itmAchat.ListSubItems(I_COL_ACHAT_DESCR).Tag = itmAvant.ListSubItems(I_COL_ACHAT_DESCR).Tag

365       itmAchat.ListSubItems(I_COL_ACHAT_DESCR).ForeColor = lColor

          'On met le nom du fabricant dans la colonne et l'ordre de la section dans le tag
370       itmAchat.SubItems(I_COL_ACHAT_MANUFACT) = itmAvant.SubItems(I_COL_ACHAT_MANUFACT)
375       itmAchat.ListSubItems(I_COL_ACHAT_MANUFACT).ForeColor = lColor

380       If bPrixSpecial = False Then
385         If optUSA.Value = True Then
390           itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixList.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
395         Else
400           If optSpain.Value = True Then
405             itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixList.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
410           Else
415             itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = Conversion(txtPrixList.Text, MODE_ARGENT, 4)
420           End If
425         End If

430         itmAchat.ListSubItems(I_COL_ACHAT_PRIX_LIST).Tag = txtPrixList.Text
       
435         itmAchat.ListSubItems(I_COL_ACHAT_PRIX_LIST).ForeColor = lColor
       
            'Escompte
440         If mskEscompte.Text <> vbNullString Then
445           itmAchat.SubItems(I_COL_ACHAT_ESCOMPTE) = Conversion(mskEscompte.Text, MODE_POURCENT)
450         Else
455           itmAchat.SubItems(I_COL_ACHAT_ESCOMPTE) = Conversion("0", MODE_POURCENT)
460         End If

465         itmAchat.ListSubItems(I_COL_ACHAT_ESCOMPTE).ForeColor = lColor

            'Prix net
470         If optUSA.Value = True Then
475           itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixNet.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
480         Else
485           If optSpain.Value = True Then
490             itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixNet.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
495           Else
500             itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(txtPrixNet.Text, MODE_ARGENT, 4)
505           End If
510         End If

515         itmAchat.ListSubItems(I_COL_ACHAT_PRIX_NET).Tag = itmAvant.ListSubItems(I_COL_ACHAT_PRIX_NET).Tag

520         itmAchat.ListSubItems(I_COL_ACHAT_PRIX_NET).ForeColor = lColor
525       Else
530         If optUSA.Value = True Then
535           itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
540         Else
545           If optSpain.Value = True Then
550             itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
555           Else
560             itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = Conversion(txtPrixSpecial.Text, MODE_ARGENT, 4)
565           End If
570         End If

575         itmAchat.ListSubItems(I_COL_ACHAT_PRIX_LIST).Tag = txtPrixSpecial.Text

580         itmAchat.ListSubItems(I_COL_ACHAT_PRIX_LIST).ForeColor = lColor

585         itmAchat.SubItems(I_COL_ACHAT_ESCOMPTE) = Conversion("0", MODE_POURCENT)

590         itmAchat.ListSubItems(I_COL_ACHAT_ESCOMPTE).ForeColor = lColor

595         If optUSA.Value = True Then
600           itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
605         Else
610           If optSpain.Value = True Then
615             itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
620           Else
625             itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(txtPrixSpecial.Text, MODE_ARGENT, 4)
630           End If
635         End If

640         itmAchat.ListSubItems(I_COL_ACHAT_PRIX_NET).Tag = itmAvant.ListSubItems(I_COL_ACHAT_PRIX_NET).Tag

645         itmAchat.ListSubItems(I_COL_ACHAT_PRIX_NET).ForeColor = lColor
650       End If

          'On met le fournisseur dans la colonne et l'id dans le tag
655       itmAchat.SubItems(I_COL_ACHAT_DISTRIB) = itmAvant.SubItems(I_COL_ACHAT_DISTRIB)
660       itmAchat.ListSubItems(I_COL_ACHAT_DISTRIB).Tag = itmAvant.ListSubItems(I_COL_ACHAT_DISTRIB).Tag

665       itmAchat.ListSubItems(I_COL_ACHAT_DISTRIB).ForeColor = lColor
          
          'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
670       itmAchat.SubItems(I_COL_ACHAT_TOTAL) = Conversion(CStr(Round(itmAchat.Text * itmAchat.SubItems(I_COL_ACHAT_PRIX_NET), 2)), MODE_ARGENT)

675       itmAchat.ListSubItems(I_COL_ACHAT_TOTAL).ForeColor = lColor

680       itmAchat.SubItems(I_COL_ACHAT_DATE_COMMANDE) = itmAvant.SubItems(I_COL_ACHAT_DATE_COMMANDE)
685       itmAchat.ListSubItems(I_COL_ACHAT_DATE_COMMANDE).ForeColor = lColor

690       itmAchat.SubItems(I_COL_ACHAT_DATE_REQUISE) = itmAvant.SubItems(I_COL_ACHAT_DATE_REQUISE)
695       itmAchat.ListSubItems(I_COL_ACHAT_DATE_REQUISE).ForeColor = lColor

700       If itmAvant.SubItems(I_COL_ACHAT_DATE_COMMANDE) <> "" Then
705         itmAvant.ListSubItems(I_COL_ACHAT_DATE_COMMANDE).ForeColor = COLOR_NOIR
710       End If

715       If itmAvant.SubItems(I_COL_ACHAT_DATE_REQUISE) <> "" Then
720         itmAvant.ListSubItems(I_COL_ACHAT_DATE_REQUISE).ForeColor = COLOR_NOIR
725       End If

730       itmAvant.ListSubItems(I_COL_ACHAT_DESCR).ForeColor = COLOR_NOIR
735       itmAvant.ListSubItems(I_COL_ACHAT_DISTRIB).ForeColor = COLOR_NOIR
740       itmAvant.ListSubItems(I_COL_ACHAT_ESCOMPTE).ForeColor = COLOR_NOIR
745       itmAvant.ListSubItems(I_COL_ACHAT_MANUFACT).ForeColor = COLOR_NOIR
750       itmAvant.ListSubItems(I_COL_ACHAT_PIECE).ForeColor = COLOR_NOIR
755       itmAvant.ListSubItems(I_COL_ACHAT_PRIX_LIST).ForeColor = COLOR_NOIR
760       itmAvant.ListSubItems(I_COL_ACHAT_PRIX_NET).ForeColor = COLOR_NOIR
765       itmAvant.ForeColor = COLOR_NOIR
770       itmAvant.ListSubItems(I_COL_ACHAT_TOTAL).ForeColor = COLOR_NOIR

          'Resélectionne le premier élément du listview
775       If lvwAchat.ListItems.count > 0 Then
780         lvwAchat.ListItems(1).Selected = True
785       End If
          
790       m_bMauvaisPrix = False

795       cmbfrs.Locked = False
800     Else
805       sPiece = lvwAchat.ListItems(CInt(fraPrixPiece.Tag)).SubItems(I_COL_ACHAT_PIECE)

810       For iCompteur = 1 To lvwAchat.ListItems.count
815         If lvwAchat.ListItems(iCompteur).SubItems(I_COL_ACHAT_PIECE) = sPiece And lvwAchat.ListItems(iCompteur).ListSubItems(I_COL_ACHAT_PIECE).ForeColor = COLOR_MAGENTA Then
820           Set itmAchat = lvwAchat.ListItems(iCompteur)

825           itmAchat.ListSubItems(I_COL_ACHAT_PIECE).ForeColor = COLOR_NOIR
830           itmAchat.ListSubItems(I_COL_ACHAT_DESCR).ForeColor = COLOR_NOIR
835           itmAchat.ListSubItems(I_COL_ACHAT_DISTRIB).ForeColor = COLOR_NOIR
840           itmAchat.ListSubItems(I_COL_ACHAT_ESCOMPTE).ForeColor = COLOR_NOIR
845           itmAchat.ListSubItems(I_COL_ACHAT_MANUFACT).ForeColor = COLOR_NOIR
850           itmAchat.ListSubItems(I_COL_ACHAT_PRIX_LIST).ForeColor = COLOR_NOIR
855           itmAchat.ListSubItems(I_COL_ACHAT_PRIX_NET).ForeColor = COLOR_NOIR
860           itmAchat.ListSubItems(I_COL_ACHAT_TOTAL).ForeColor = COLOR_NOIR
865           itmAchat.ForeColor = COLOR_NOIR

870           Call lvwAchat.Refresh
  
875           If bPrixSpecial = False Then
                'Prix listé
880             If optUSA.Value = True Then
885               itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixList.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
890             Else
895               If optSpain.Value = True Then
900                 itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixList.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
905               Else
910                 itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = Conversion(txtPrixList.Text, MODE_ARGENT, 4)
915               End If
920             End If

925             itmAchat.ListSubItems(I_COL_ACHAT_PRIX_LIST).Tag = txtPrixList.Text
        
                'Escompte
930             If mskEscompte.Text <> vbNullString Then
935               itmAchat.SubItems(I_COL_ACHAT_ESCOMPTE) = Conversion(mskEscompte.Text, MODE_POURCENT)
940             Else
945               itmAchat.SubItems(I_COL_ACHAT_ESCOMPTE) = Conversion("0", MODE_POURCENT)
950             End If

                'Prix net
955             If optUSA.Value = True Then
960               itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixNet.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
965             Else
970               If optSpain.Value = True Then
975                 itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixNet.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
980               Else
985                 itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(txtPrixNet.Text, MODE_ARGENT, 4)
990               End If
995             End If
1000          Else
1005            If optUSA.Value = True Then
1010              itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
1015            Else
1020              If optSpain.Value = True Then
1025                itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
1030              Else
1035                itmAchat.SubItems(I_COL_ACHAT_PRIX_LIST) = Conversion(txtPrixSpecial.Text, MODE_ARGENT, 4)
1040              End If
1045            End If

1050            itmAchat.ListSubItems(I_COL_ACHAT_PRIX_LIST).Tag = txtPrixSpecial.Text
         
1055            itmAchat.SubItems(I_COL_ACHAT_ESCOMPTE) = Conversion("0", MODE_POURCENT)

1060            If optUSA.Value = True Then
1065              itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxUSA), 4)), MODE_ARGENT, 4)
1070            Else
1075              If optSpain.Value = True Then
1080                itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(CStr(Round(CDbl(txtPrixSpecial.Text) / CDbl(sTauxSPA), 4)), MODE_ARGENT, 4)
1085              Else
1090                itmAchat.SubItems(I_COL_ACHAT_PRIX_NET) = Conversion(txtPrixSpecial.Text, MODE_ARGENT, 4)
1095              End If
1100            End If
1105          End If

              'On met le fournisseur dans la colonne et l'id dans le tag
1110          itmAchat.SubItems(I_COL_ACHAT_DISTRIB) = cmbfrs.LIST(cmbfrs.ListIndex)
  
1115          itmAchat.ListSubItems(I_COL_ACHAT_DISTRIB).Tag = cmbfrs.ItemData(cmbfrs.ListIndex)
         
              'Pour le prix total, il faut faire la quantité * prix net * pourcentage de profit
1120          itmAchat.SubItems(I_COL_ACHAT_TOTAL) = Conversion(CStr(Round(Replace(itmAchat.Text, "*", "") * itmAchat.SubItems(I_COL_ACHAT_PRIX_NET), 2)), MODE_ARGENT)

1125          If optUSA.Value = True Then
1130            itmAchat.ListSubItems(I_COL_ACHAT_TOTAL).Tag = "USA"
1135          Else
1140            If optSpain.Value = True Then
1145              itmAchat.ListSubItems(I_COL_ACHAT_TOTAL).Tag = "SPA"
1150            Else
1155              itmAchat.ListSubItems(I_COL_ACHAT_TOTAL).Tag = "CAN"
1160            End If
1165          End If

1170        End If
1175      Next
1180    End If

1185    Call ModifierPrixCatalogue

1190    fraPrixPiece.Visible = False

1195    Call CalculerPrix

1200    Exit Sub

AfficherErreur:

1205    woups "frmAchat", "cmdOKPrix_Click", Err, Erl
End Sub

Private Sub RemplirComboFournisseur()

5       On Error GoTo AfficherErreur

10      Dim rstFRS As ADODB.Recordset

        'Il faut vider le combo avant de le remplir
15      Call cmbfrs.Clear

20      Set rstFRS = New ADODB.Recordset

25      Call rstFRS.Open("SELECT GRB_PiecesFRS.*, GRB_Fournisseur.NomFournisseur FROM GRB_PiecesFRS INNER JOIN GRB_Fournisseur ON GRB_PiecesFRS.IDFRS = GRB_Fournisseur.IDFRS WHERE PIECE = '" & Replace(lvwAchat.SelectedItem.SubItems(I_COL_ACHAT_PIECE), "'", "''") & "' ORDER BY NomFournisseur", g_connData, adOpenDynamic, adLockOptimistic)

        'Tant que ce n'est pas la fin des enregistrements
30      Do While Not rstFRS.EOF
35        Call cmbfrs.AddItem(rstFRS.Fields("NomFournisseur"))

40        cmbfrs.ItemData(cmbfrs.newIndex) = rstFRS.Fields("IDFRS")

45        Call rstFRS.MoveNext
50      Loop

55      Exit Sub

AfficherErreur:

60      woups "frmAchat", "RemplirComboFournisseur", Err, Erl
End Sub

Private Sub mnuDateRequise_Click()

5       On Error GoTo AfficherErreur

10      If Trim$(lvwAchat.SelectedItem.SubItems(I_COL_ACHAT_DATE_REQUISE)) = "" Then
15        mvwDateRequise.Year = Year(Date)
20        mvwDateRequise.Month = Month(Date)
25        mvwDateRequise.Day = Day(Date)
30      Else
35        mvwDateRequise.Year = Left$(lvwAchat.SelectedItem.SubItems(I_COL_ACHAT_DATE_REQUISE), 4)
40        mvwDateRequise.Month = Mid$(lvwAchat.SelectedItem.SubItems(I_COL_ACHAT_DATE_REQUISE), 6, 2)
45        mvwDateRequise.Day = Right$(lvwAchat.SelectedItem.SubItems(I_COL_ACHAT_DATE_REQUISE), 2)
50      End If

55      fraDateRequise.Top = lvwAchat.Top

60      fraDateRequise.Visible = True

65      Exit Sub

AfficherErreur:

70      woups "frmAchat", "mnuDateRequise_Click", Err, Erl
End Sub

Private Sub mvwDateRequise_GotFocus()
  
5       On Error GoTo AfficherErreur

10      m_bMonthViewHasFocus = True

15      Exit Sub

AfficherErreur:

20      woups "frmAchat", "mvwDateRequise_GotFocus", Err, Erl
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

60      woups "frmAchat", "txtPrixList_LostFocus", Err, Erl
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

45      woups "frmAchat", "txtPrixNet_Change", Err, Erl
End Sub

Private Sub txtPrixNet_GotFocus()

5       On Error GoTo AfficherErreur

        'Si le prix net prend le focus
10      Call CalculerPrixNet

15      Exit Sub

AfficherErreur:

20      woups "frmAchat", "txtPrixNet_GotFocus", Err, Erl
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

105     woups "frmAchat", "CalculerPrixNet", Err, Erl
End Sub

Private Sub txtPrixNet_LostFocus()

5       On Error GoTo AfficherErreur

10      txtPrixNet.Text = Replace(txtPrixNet.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmAchat", "txtPrixNet_LostFocus", Err, Erl
End Sub

Private Sub ViderChamps_frs()

5       On Error GoTo AfficherErreur

        'Vide les champs pieces
10      txtPrixList.Text = vbNullString
15      mskEscompte.Text = vbNullString
20      txtPrixNet.Text = vbNullString
  
25      optCAN.Value = True

30      Call AfficherDrapeau

35      Exit Sub

AfficherErreur:

40      woups "frmAchat", "ViderChamps_frs", Err, Erl
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
        
        'Ouverture du recordset
110     Call rstPrix.Open("SELECT * FROM GRB_PiecesFRS WHERE PIECE = '" & Replace(lvwAchat.ListItems(CInt(fraPrixPiece.Tag)).SubItems(I_COL_ACHAT_PIECE), "'", "''") & "' AND IDFRS = " & cmbfrs.ItemData(cmbfrs.ListIndex), g_connData, adOpenDynamic, adLockOptimistic)
  
115     rstPrix.Fields("PRIX_LIST") = dblPrixList
120     rstPrix.Fields("ESCOMPTE") = dblEscompte
125     rstPrix.Fields("PRIX_NET") = dblPrixNet
130     rstPrix.Fields("DATE") = ConvertDate(Date)
135     rstPrix.Fields("ENTRER_PAR") = g_sInitiale
  
140     If optCAN.Value = True Then
145       rstPrix.Fields("DeviseMonétaire") = "CAN"
150     Else
155       If optUSA.Value = True Then
160         rstPrix.Fields("DeviseMonétaire") = "USA"
165       Else
170         rstPrix.Fields("DeviseMonétaire") = "SPA"
175       End If
180     End If
    
185     If m_eCatalogue = ELECTRIQUE Then
190       rstPrix.Fields("Type") = "E"
195     Else
200       rstPrix.Fields("Type") = "M"
205     End If
  
210     Call rstPrix.Update
  
215     Call rstPrix.Close
220     Set rstPrix = Nothing
        
225     Exit Sub

AfficherErreur:

230     woups "frmAchat", "ModifierPrixCatalogue", Err, Erl
End Sub

Private Sub optCAN_Click()

5       On Error GoTo AfficherErreur

        'Dépendant la devise, affiche le drapeau
10      Call AfficherDrapeau

15      Exit Sub

AfficherErreur:

20      woups "frmAchat", "optCAN_Click", Err, Erl
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

85     Exit Sub

AfficherErreur:

90      woups "frmAchat", "AfficherDrapeau", Err, Erl
End Sub

Private Sub optSpain_Click()

5       On Error GoTo AfficherErreur

        'dependant la devise, affiche le drapeau
10      Call AfficherDrapeau

15      Exit Sub

AfficherErreur:

20      woups "frmAchat", "optSpain_Click", Err, Erl
End Sub

Private Sub optUSA_Click()

5       On Error GoTo AfficherErreur

        'dependant la devise, affiche le drapeau
10      Call AfficherDrapeau

15      Exit Sub

AfficherErreur:

20      woups "frmAchat", "optUSA_Click", Err, Erl
End Sub

Private Sub mskEscompte_GotFocus()

5       On Error GoTo AfficherErreur

        'Quand le maskEdit prend le focus, on set le masque
10      If mskEscompte.Enabled = True Then
15        mskEscompte.mask = "0,####"
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmAchat", "mskEscompte_GotFocus", Err, Erl
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

40      woups "frmAchat", "mskEscompte_LostFocus", Err, Erl
End Sub

Public Sub Commande()

5       On Error GoTo AfficherErreur

10      Dim rstPiece        As ADODB.Recordset
15      Dim rstBCPiece      As ADODB.Recordset
20      Dim rstBC           As ADODB.Recordset
25      Dim rstFRS          As ADODB.Recordset
30      Dim iIDFRS          As Integer
35      Dim sFRS            As String
40      Dim sNoBC           As String
45      Dim sWhere          As String
50      Dim sWherePiece     As String
55      Dim sWhereNoLigne   As String
60      Dim sDateRequise    As String
65      Dim sNoLigne        As String
70      Dim bPremier        As Boolean
75      Dim bPremierNoLigne As Boolean

80      sFRS = DR_Commande.Sections("Section2").Controls("lblFournisseur").Caption
85      sNoBC = DR_Commande.Sections("Section2").Controls("lblNoBC").Caption

90      Set rstBC = New ADODB.Recordset
95      Set rstFRS = New ADODB.Recordset
100     Set rstPiece = New ADODB.Recordset
105     Set rstBCPiece = New ADODB.Recordset

110     Call rstBC.Open("SELECT * FROM GRB_BonsCommandes WHERE NoBonCommande = '" & sNoBC & "'", g_connData, adOpenDynamic, adLockOptimistic)

115     Do While Not rstBC.EOF
120       Call rstFRS.Open("SELECT IDFRS, NomFournisseur FROM GRB_Fournisseur WHERE IDFRS = " & rstBC.Fields("NoFournisseur"), g_connData, adOpenDynamic, adLockOptimistic)

125       If rstFRS.Fields("NomFournisseur") = sFRS Then
130         iIDFRS = rstFRS.Fields("IDFRS")

135         sDateRequise = rstBC.Fields("DateRequise")

140         Call rstFRS.Close

145         Exit Do
150       End If

155       Call rstFRS.Close

160       Call rstBC.MoveNext
165     Loop

170     Call rstBC.Close
175     Set rstBC = Nothing

180     Set rstFRS = Nothing
        
        'Ouverture du recordset du Bon de commande pour savoir quelles pièces
        'ont été commandées
185     Call rstBCPiece.Open("SELECT NoItem, NuméroLigne FROM GRB_BonsCommandes_Pieces WHERE NoFournisseur = " & iIDFRS & " AND NoBonCommande = '" & sNoBC & "'", g_connData, adOpenDynamic, adLockOptimistic)

        'Tant que ce n'est pas la fin des enregistrements
190     sWhere = "(IDAchat = '" & Left$(txtNoAchat.Text, 9) & "' AND IndexAchat = " & Int(Right$(txtNoAchat.Text, 3)) & ")"

195     sWherePiece = "PIECE In ("
200     sWhereNoLigne = "NuméroLigne In ("
        
205     bPremier = True
        
210     Do While Not rstBCPiece.EOF
215       If Not IsNull(rstBCPiece.Fields("NoItem")) Then
220         sNoLigne = rstBCPiece.Fields("NuméroLigne")

225         If bPremier = True Then
230           If InStr(1, sNoLigne, ",") = 0 Then
235             sWherePiece = sWherePiece & "'" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
240             sWhereNoLigne = sWhereNoLigne & rstBCPiece.Fields("NuméroLigne")
245           Else
250             bPremierNoLigne = True

255             Do While InStr(1, sNoLigne, ",") > 0
260               If bPremierNoLigne = True Then
265                 sWherePiece = sWherePiece & "'" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
270                 sWhereNoLigne = sWhereNoLigne & Left$(sNoLigne, InStr(1, sNoLigne, ",") - 1)

275                 bPremierNoLigne = False
280               Else
285                 sWherePiece = sWherePiece & ", '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
290                 sWhereNoLigne = sWhereNoLigne & ", " & Left$(sNoLigne, InStr(1, sNoLigne, ",") - 1)
295               End If

300               sNoLigne = Right$(sNoLigne, Len(sNoLigne) - (InStr(1, sNoLigne, ",") + 1))
305             Loop

310             If Trim$(sNoLigne) <> "" Then
315               sWherePiece = sWherePiece & ", '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
320               sWhereNoLigne = sWhereNoLigne & ", " & sNoLigne
325             End If
330           End If

335           bPremier = False
340         Else
345           If InStr(1, sNoLigne, ",") = 0 Then
350             sWherePiece = sWherePiece & ", '" & rstBCPiece.Fields("NoItem") & "'"
355             sWhereNoLigne = sWhereNoLigne & ", " & rstBCPiece.Fields("NuméroLigne")
360           Else
365             Do While InStr(1, sNoLigne, ",") > 0
370               sWherePiece = sWherePiece & ", '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
375               sWhereNoLigne = sWhereNoLigne & ", " & Left$(sNoLigne, InStr(1, sNoLigne, ",") - 1)

380               sNoLigne = Right$(sNoLigne, Len(sNoLigne) - (InStr(1, sNoLigne, ",") + 1))
385             Loop

390             If Trim$(sNoLigne) <> "" Then
395               sWherePiece = sWherePiece & ", '" & Replace(rstBCPiece.Fields("NoItem"), "'", "''") & "'"
400               sWhereNoLigne = sWhereNoLigne & ", " & sNoLigne
405             End If
410           End If
415         End If
420       End If
   
425       Call rstBCPiece.MoveNext
430     Loop

435     sWherePiece = sWherePiece & ")"
440     sWhereNoLigne = sWhereNoLigne & ")"

445     sWhere = sWhere & " AND " & sWherePiece & " AND " & sWhereNoLigne
  
450     Call rstBCPiece.Close
455     Set rstBCPiece = Nothing

460     Call rstPiece.Open("SELECT * FROM GRB_Achat_Pieces WHERE " & sWhere, g_connData, adOpenDynamic, adLockOptimistic)
  
465     Do While Not rstPiece.EOF
470       rstPiece.Fields("Commandé") = True

475       rstPiece.Fields("DateCommande") = ConvertDate(Date)

480       rstPiece.Fields("DateRequise") = sDateRequise
    
485       Call rstPiece.Update
    
490       Call rstPiece.MoveNext
495     Loop
  
500     Call rstPiece.Close
505     Set rstPiece = Nothing
          
510     Call RemplirListViewAchat

515     Exit Sub

AfficherErreur:

520     woups "frmAchat", "Commande", Err, Erl
End Sub

Private Sub cmdRetour_Click()

5       On Error GoTo AfficherErreur

10      Dim sIDAchat    As String
15      Dim iIndexAchat As Integer
20      Dim rstAchat    As ADODB.Recordset

25      If cmbNoAchat.ListCount > 0 Then
30        sIDAchat = Trim$(Left$(txtNoAchat.Text, InStr(1, txtNoAchat.Text, "-") + 2))
35        iIndexAchat = CInt(Trim$(Right$(txtNoAchat.Text, Len(txtNoAchat.Text) - InStrRev(txtNoAchat.Text, "-"))))

40        Set rstAchat = New ADODB.Recordset

45        Call rstAchat.Open("SELECT Modification, Par FROM GRB_Achat WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat, g_connData, adOpenDynamic, adLockOptimistic)

50        If rstAchat.Fields("Modification") = False Then
55          Call rstAchat.Close
60          Set rstAchat = Nothing

65          Screen.MousePointer = vbHourglass
  
70          iIndexAchat = CInt(Right$(txtNoAchat.Text, 3))

75          Call frmRetourMarchandise.AfficherAchat(sIDAchat, iIndexAchat, g_sUserID)

80          Call cmbNoAchat_Click

85          Screen.MousePointer = vbDefault
90        Else
95          Call MsgBox("Cet achat est en modification par " & rstAchat.Fields("Par") & "!", vbOKOnly, "Erreur")

100         Call rstAchat.Close
105         Set rstAchat = Nothing
110       End If
115     End If

120     Exit Sub

AfficherErreur:

125     woups "frmAchat", "cmdRetour_Click", Err, Erl
End Sub

Private Sub OuvrirAchat(ByVal bOuvrir As Boolean)
        'Remplit ou vide les champs Modification et Par
5       On Error GoTo AfficherErreur

10      Dim rstAchat    As ADODB.Recordset
15      Dim sIDAchat    As String
20      Dim iIndexAchat As Integer

25      sIDAchat = Trim$(Left$(txtNoAchat.Text, InStr(1, txtNoAchat.Text, "-") + 2))
30      iIndexAchat = CInt(Trim$(Right$(txtNoAchat.Text, Len(txtNoAchat.Text) - InStrRev(txtNoAchat.Text, "-"))))

35      Set rstAchat = New ADODB.Recordset

40      rstAchat.CursorLocation = adUseServer

45      Call rstAchat.Open("SELECT Modification, Par FROM GRB_Achat WHERE IDAchat = '" & sIDAchat & "' AND IndexAchat = " & iIndexAchat, g_connData, adOpenDynamic, adLockOptimistic)

50      Do While Not rstAchat.EOF
55        If bOuvrir = True Then
60          rstAchat.Fields("Modification") = True
65          rstAchat.Fields("Par") = g_sEmploye
70        Else
75          rstAchat.Fields("Modification") = False
80          rstAchat.Fields("Par") = ""
85        End If

90        Call rstAchat.Update

95        Call rstAchat.MoveNext
100     Loop

105     Call rstAchat.Close
110     Set rstAchat = Nothing

115     Exit Sub

AfficherErreur:

120     woups "frmAchat", "OuvrirAchat", Err, Erl
End Sub

Private Sub lvwPieceTrouve_DblClick()

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer

15      m_bRecherchePiece = True
20      m_bInventaire = False

        'On affiche lvwFournisseur selon la pièce choisie
        'Rempli les fournisseurs de la pièce choisie
25      Call AfficherListeFournisseurs
  
        'si le listview n'est pas vide
30      If lvwfournisseur.ListItems.count = 1 Then
35        If MsgBox("Il n'y a aucun fournisseur pour cette pièce!" & vbNewLine & "Voulez-vous en ajouter?", vbYesNo, "Erreur") = vbYes Then
40          Screen.MousePointer = vbHourglass
      
            'On ouvre le catalogue sur cet enregistrement
45          If m_eCatalogue = ELECTRIQUE Then
50            Call FrmCatalogueElec.AfficherForm(lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_CATEGORIE), lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_MANUFACT), lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM))
55          Else
60            Call FrmCatalogueMec.AfficherForm(lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_CATEGORIE), lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_MANUFACT), lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM))
65          End If

70          Screen.MousePointer = vbDefault
      
            'On rappelle la méthode
75          Call AfficherListeFournisseurs
80        End If
85      End If

90      Exit Sub

AfficherErreur:

95      woups "frmAchat", "lvwPieceTrouve_DblClick", Err, Erl
End Sub

Private Sub cmdOKPieceTrouve_Click()

5       On Error GoTo AfficherErreur

10      m_bRecherchePiece = True
15      m_bInventaire = False

        'On affiche lvwFournisseur selon la pièce choisie
        'Rempli les fournisseurs de la pièce choisie
20      Call AfficherListeFournisseurs
  
        'si le listview n'est pas vide
25      If lvwfournisseur.ListItems.count = 1 Then
30        If MsgBox("Il n'y a aucun fournisseur pour cette pièce!" & vbNewLine & "Voulez-vous en ajouter?", vbYesNo, "Erreur") = vbYes Then
35          Screen.MousePointer = vbHourglass
      
            'On ouvre le catalogue sur cet enregistrement
40          If m_eCatalogue = ELECTRIQUE Then
45            Call FrmCatalogueElec.AfficherForm(lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_CATEGORIE), lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_MANUFACT), lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM))
50          Else
55            Call FrmCatalogueMec.AfficherForm(lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_CATEGORIE), lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_MANUFACT), lvwPieceTrouve.SelectedItem.SubItems(I_COL_RECH_NO_ITEM))
60          End If
      
65          Screen.MousePointer = vbDefault
      
            'On rappelle la méthode
70          Call AfficherListeFournisseurs
75        End If
80      End If

85      Exit Sub

AfficherErreur:

90      woups "frmAchat", "cmdOKPieceTrouve_Click", Err, Erl
End Sub

Private Sub cmdAnnulerPieceTrouve_Click()

5       On Error GoTo AfficherErreur

10      fraPieceTrouve.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmAchat", "cmdAnnulerPieceTrouve", Err, Erl
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

120     If m_eCatalogue = ELECTRIQUE Then
125       If iIndexColumn = I_COL_PIECES_NO_ITEM Then
130         Call rstPiece.Open("SELECT * FROM GRB_CatalogueElec WHERE INSTR(1, " & sChamps & ", '" & sRecherche & "') > 0", g_connData, adOpenDynamic, adLockOptimistic)
135       Else
140         Call rstPiece.Open("SELECT * FROM GRB_CatalogueElec WHERE INSTR(1, " & sChamps & ", '" & sTexte & "') > 0", g_connData, adOpenDynamic, adLockOptimistic)
145       End If
150     Else
155       If iIndexColumn = I_COL_PIECES_NO_ITEM Then
160         Call rstPiece.Open("SELECT * FROM GRB_CatalogueMec WHERE INSTR(1, " & sChamps & ", '" & sRecherche & "') > 0", g_connData, adOpenDynamic, adLockOptimistic)
165       Else
170         Call rstPiece.Open("SELECT * FROM GRB_CatalogueMec WHERE INSTR(1, " & sChamps & ", '" & sTexte & "') > 0", g_connData, adOpenDynamic, adLockOptimistic)
175       End If
180     End If

        'Pour chaque enregistrement
185     Do While Not rstPiece.EOF
          'On ajoute dans le ListView
190       Set itmPiece = lvwPieceTrouve.ListItems.Add

195       If m_eCatalogue = ELECTRIQUE Then
200         If Not IsNull(rstPiece.Fields("TEMPS")) Then
205           itmPiece.Tag = rstPiece.Fields("TEMPS")
210         Else
215           itmPiece.Tag = vbNullString
220         End If
225       End If

230       If Not IsNull(rstPiece.Fields("PIECE_GRB")) Then
235         itmPiece.Text = rstPiece.Fields("PIECE_GRB")
240       Else
245         itmPiece.Text = ""
250       End If

255       itmPiece.SubItems(I_COL_RECH_NO_ITEM) = rstPiece.Fields("PIECE")
260       itmPiece.SubItems(I_COL_RECH_CATEGORIE) = cmbCategorie.LIST(iCompteur)

265       If Not IsNull(rstPiece.Fields("FABRICANT")) Then
270         itmPiece.SubItems(I_COL_RECH_MANUFACT) = rstPiece.Fields("FABRICANT")
275       Else
280         itmPiece.SubItems(I_COL_RECH_MANUFACT) = ""
285       End If

290       If Not IsNull(rstPiece.Fields("DESC_EN")) Then
295         itmPiece.SubItems(I_COL_RECH_DESCR_EN) = rstPiece.Fields("DESC_EN")
300       Else
305         itmPiece.SubItems(I_COL_RECH_DESCR_EN) = ""
310       End If

315       If Not IsNull(rstPiece.Fields("DESC_FR")) Then
320         itmPiece.SubItems(I_COL_RECH_DESCR_FR) = rstPiece.Fields("DESC_FR")
325       Else
330         itmPiece.SubItems(I_COL_RECH_DESCR_FR) = ""
335       End If

340       Call rstPiece.MoveNext
345     Loop

350     Call rstPiece.Close
355     Set rstPiece = Nothing

360     Exit Sub

AfficherErreur:

365     woups "frmAchat", "RemplirListViewRecherche", Err, Erl
End Sub

Private Sub cmdMaterielInutile_Click()

5       On Error GoTo AfficherErreur

10      Dim itmAchat As ListItem

15      If lvwAchat.ListItems.count > 0 Then
20        Set itmAchat = lvwAchat.SelectedItem
          
          'Si la quantité est plus grande que 0
25        If CDbl(itmAchat.Text) > 0 Then
30          m_bPieceInutile = True
35          m_bRecherchePiece = False

40          Call AfficherListeFournisseurs

45          If lvwfournisseur.ListItems.count = 0 Then
50            Call MsgBox("Il n'y a aucun fournisseur pour cette pièce!", vbOKOnly, "Erreur")
 
55            Exit Sub
60          Else
65            frafournisseur.Visible = True
70          End If
75        Else
80          Call MsgBox("La quantité est déjà dans le négatif!", vbOKOnly, "Erreur")
85        End If
90      End If

95      Exit Sub

AfficherErreur:

100     woups "frmAchat", "cmdMaterielInutile_Click", Err, Erl
End Sub

Private Sub cmdMauvaisPrix_Click()

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer
15      Dim itmAchat  As ListItem

20      If lvwAchat.ListItems.count > 0 Then
25        Set itmAchat = lvwAchat.SelectedItem
          
          'Si la quantité est plus grande que 0
30        If CDbl(itmAchat.Text) > 0 Then
35          Call ViderChamps_frs

40          Call RemplirComboFournisseur

45          For iCompteur = 0 To cmbfrs.ListCount - 1
50            If cmbfrs.ItemData(iCompteur) = itmAchat.ListSubItems(I_COL_ACHAT_DISTRIB).Tag Then
55              cmbfrs.ListIndex = iCompteur

60              Exit For
65            End If
70          Next

75          cmbfrs.Locked = True

80          fraPrixPiece.Tag = itmAchat.Index

85          m_bMauvaisPrix = True

90          fraPrixPiece.Visible = True
95        Else
100         Call MsgBox("La quantité est déjà dans le négatif!", vbOKOnly, "Erreur")
105       End If
110     End If

115     Exit Sub

AfficherErreur:

120     woups "frmAchat", "cmdMauvaisPrix_Click", Err, Erl
End Sub

Private Sub cmdReception_Click()

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer
15      Dim bOuvert   As Boolean

20      If m_eCatalogue = ELECTRIQUE Then
25        For iCompteur = 0 To Forms.count - 1
30          If Forms(iCompteur).Name = "FrmReceptionElec" Then
35            bOuvert = True

40            Exit For
45          End If
50        Next

55        If bOuvert = True Then
60          Call Unload(FrmReceptionElec)
65        End If

70        Call FrmReceptionElec.AfficherAchat(g_sUserID, txtNoAchat.Text)

75        Call RemplirListViewAchat
80      Else
85        For iCompteur = 0 To Forms.count - 1
90          If Forms(iCompteur).Name = "FrmReceptionMec" Then
95            bOuvert = True

100           Exit For
105         End If
110       Next

115       If bOuvert = True Then
120         Call Unload(FrmReceptionMec)
125       End If

130       Call FrmReceptionMec.AfficherAchat(g_sUserID, txtNoAchat.Text)

135       Call RemplirListViewAchat
140     End If

145     Exit Sub

AfficherErreur:

150     woups "frmAchat", "cmdReception_Click", Err, Erl
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

60      woups "frmAchat", "txtPrixSpecial_Change", Err, Erl
End Sub

Private Sub txtPrixSpecial_LostFocus()

5       On Error GoTo AfficherErreur

10      txtPrixSpecial.Text = Replace(txtPrixSpecial.Text, ".", ",")

15      Exit Sub

AfficherErreur:

20      woups "frmAchat", "txtPrixSpecial_LostFocus", Err, Erl
End Sub
