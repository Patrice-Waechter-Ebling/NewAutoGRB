VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPunch 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Punch"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmPunch.frx":0000
   ScaleHeight     =   8760
   ScaleWidth      =   13980
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraJour 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   4800
      TabIndex        =   0
      Top             =   600
      Width           =   9135
      Begin VB.CommandButton cmdPunchMultiple 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Punch multiple"
         Height          =   375
         Left            =   1800
         TabIndex        =   53
         Top             =   4800
         Width           =   1215
      End
      Begin MSComctlLib.ListView lvwJour 
         Height          =   4695
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   8281
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nom"
            Object.Width           =   926
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Projet"
            Object.Width           =   1588
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Début"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fin"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Client"
            Object.Width           =   3889
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Commentaire"
            Object.Width           =   2752
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "KM"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton cmdPunchOut 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Punch out"
         Height          =   375
         Left            =   7320
         TabIndex        =   5
         Top             =   4800
         Width           =   975
      End
      Begin VB.CommandButton cmdModifierPunchOut 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Modifier punch out"
         Height          =   375
         Left            =   4680
         TabIndex        =   3
         Top             =   4800
         Width           =   1575
      End
      Begin VB.CommandButton cmdPunchIn 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Punch in"
         Height          =   375
         Left            =   6360
         TabIndex        =   4
         Top             =   4800
         Width           =   855
      End
      Begin VB.CommandButton cmdModifierPunchIn 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Modifier punch in"
         Height          =   375
         Left            =   3120
         TabIndex        =   2
         Top             =   4800
         Width           =   1455
      End
   End
   Begin VB.Frame fraPunchMultiple 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   4800
      TabIndex        =   54
      Top             =   600
      Visible         =   0   'False
      Width           =   8415
      Begin VB.ComboBox cmbPMType 
         Height          =   315
         ItemData        =   "frmPunch.frx":2F0D
         Left            =   2760
         List            =   "frmPunch.frx":2F0F
         Style           =   2  'Dropdown List
         TabIndex        =   85
         Top             =   3240
         Width           =   5415
      End
      Begin VB.OptionButton optPMHeureDiner 
         Caption         =   "1 heure"
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   5760
         TabIndex        =   75
         Top             =   4440
         Width           =   1215
      End
      Begin VB.OptionButton optPMHeureDiner 
         Caption         =   "30 minutes"
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   5760
         TabIndex        =   74
         Top             =   4200
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmdBrowseCommentPM 
         Caption         =   "Choisir un commentaire"
         Height          =   375
         Left            =   3240
         TabIndex        =   70
         Top             =   3840
         Width           =   1935
      End
      Begin VB.CommandButton cmdPMOK 
         BackColor       =   &H00C0C0C0&
         Caption         =   "OK"
         Height          =   375
         Left            =   7200
         TabIndex        =   61
         Top             =   4800
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         Caption         =   "Heure"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   59
         Top             =   2400
         Width           =   2055
         Begin MSMask.MaskEdBox mskPMHeureFin 
            Height          =   255
            Left            =   840
            TabIndex        =   60
            Top             =   600
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskPMHeureDebut 
            Height          =   255
            Left            =   840
            TabIndex        =   66
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.Label Label10 
            Caption         =   "Fin :"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label8 
            Caption         =   "Début :"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox txtPMCommentaire 
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   58
         Top             =   4320
         Width           =   5055
      End
      Begin VB.CommandButton cmdPMAnnuler 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Annuler"
         Height          =   375
         Left            =   6000
         TabIndex        =   57
         Top             =   4800
         Width           =   1095
      End
      Begin VB.TextBox txtPMClient 
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   2640
         Width           =   5415
      End
      Begin VB.CheckBox chkPMDiner 
         Caption         =   "Heure de dîner"
         Height          =   255
         Left            =   5280
         TabIndex        =   55
         Top             =   3840
         Width           =   1455
      End
      Begin MSComctlLib.ListView lvwEmployes 
         Height          =   1575
         Left            =   0
         TabIndex        =   65
         Top             =   0
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   2778
         View            =   3
         LabelEdit       =   1
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Employé"
            Object.Width           =   11112
         EndProperty
      End
      Begin MSMask.MaskEdBox mskPMNoProjet 
         Height          =   255
         Left            =   1320
         TabIndex        =   69
         Top             =   1800
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Mask            =   "#####-##"
         PromptChar      =   "_"
      End
      Begin VB.PictureBox picTypePunch 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   2640
         ScaleHeight     =   735
         ScaleWidth      =   1335
         TabIndex        =   76
         Top             =   1680
         Width           =   1335
         Begin VB.OptionButton optPMTypePunch 
            Caption         =   "Mécanique"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   78
            Top             =   300
            Width           =   1095
         End
         Begin VB.OptionButton optPMTypePunch 
            Caption         =   "Électrique"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   77
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.Label lblPMTypePunch 
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4320
         TabIndex        =   88
         Top             =   1680
         Width           =   3855
      End
      Begin VB.Label lblPMType 
         Caption         =   "Type"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2760
         TabIndex        =   84
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "Client"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2760
         TabIndex        =   64
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label lblPMPrefixe 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   82
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   "No. Projet"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   63
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Commentaires"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   4080
         Width           =   1095
      End
   End
   Begin VB.ComboBox cmbHeureSemaine 
      Height          =   315
      Left            =   8640
      TabIndex        =   51
      Text            =   "Combo1"
      Top             =   240
      Width           =   2535
   End
   Begin VB.Frame fraSemaine 
      BackColor       =   &H00404040&
      Height          =   2655
      Left            =   120
      TabIndex        =   28
      Top             =   6000
      Width           =   12985
      Begin MSComctlLib.ListView lvwJourSemaine 
         Height          =   1900
         Index           =   1
         Left            =   0
         TabIndex        =   43
         Top             =   720
         Width           =   1855
         _ExtentX        =   3281
         _ExtentY        =   3360
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "nom"
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "heure"
            Object.Width           =   1774
         EndProperty
      End
      Begin MSComctlLib.ListView lvwJourSemaine 
         Height          =   1900
         Index           =   2
         Left            =   1855
         TabIndex        =   44
         Top             =   720
         Width           =   1855
         _ExtentX        =   3281
         _ExtentY        =   3360
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "nom"
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "heure"
            Object.Width           =   1774
         EndProperty
      End
      Begin MSComctlLib.ListView lvwJourSemaine 
         Height          =   1900
         Index           =   3
         Left            =   3710
         TabIndex        =   45
         Top             =   720
         Width           =   1855
         _ExtentX        =   3281
         _ExtentY        =   3360
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "nom"
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "heure"
            Object.Width           =   1774
         EndProperty
      End
      Begin MSComctlLib.ListView lvwJourSemaine 
         Height          =   1900
         Index           =   4
         Left            =   5565
         TabIndex        =   46
         Top             =   720
         Width           =   1855
         _ExtentX        =   3281
         _ExtentY        =   3360
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "nom"
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "heure"
            Object.Width           =   1774
         EndProperty
      End
      Begin MSComctlLib.ListView lvwJourSemaine 
         Height          =   1905
         Index           =   5
         Left            =   7420
         TabIndex        =   47
         Top             =   720
         Width           =   1855
         _ExtentX        =   3281
         _ExtentY        =   3360
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "nom"
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "heure"
            Object.Width           =   1774
         EndProperty
      End
      Begin MSComctlLib.ListView lvwJourSemaine 
         Height          =   1905
         Index           =   6
         Left            =   9275
         TabIndex        =   48
         Top             =   720
         Width           =   1855
         _ExtentX        =   3281
         _ExtentY        =   3360
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "nom"
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "heure"
            Object.Width           =   1774
         EndProperty
      End
      Begin MSComctlLib.ListView lvwJourSemaine 
         Height          =   1905
         Index           =   7
         Left            =   11130
         TabIndex        =   49
         Top             =   720
         Width           =   1855
         _ExtentX        =   3281
         _ExtentY        =   3360
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "nom"
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "heure"
            Object.Width           =   1774
         EndProperty
      End
      Begin VB.Label lblNomJour 
         BackStyle       =   0  'Transparent
         Caption         =   "Dim"
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
         TabIndex        =   30
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblNomJour 
         BackStyle       =   0  'Transparent
         Caption         =   "Lun"
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
         Left            =   1960
         TabIndex        =   32
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblNomJour 
         BackStyle       =   0  'Transparent
         Caption         =   "Mar"
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
         Left            =   3815
         TabIndex        =   34
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblNomJour 
         BackStyle       =   0  'Transparent
         Caption         =   "Mer"
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
         Left            =   5665
         TabIndex        =   35
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblNomJour 
         BackStyle       =   0  'Transparent
         Caption         =   "Jeu"
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
         Left            =   7520
         TabIndex        =   38
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblNomJour 
         BackStyle       =   0  'Transparent
         Caption         =   "Ven"
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
         Left            =   9375
         TabIndex        =   39
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblNomJour 
         BackStyle       =   0  'Transparent
         Caption         =   "Sam"
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
         Left            =   11230
         TabIndex        =   41
         Top             =   360
         Width           =   480
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   3710
         X2              =   3710
         Y1              =   360
         Y2              =   2640
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   5565
         X2              =   5565
         Y1              =   360
         Y2              =   2640
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   7420
         X2              =   7420
         Y1              =   360
         Y2              =   2640
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         X1              =   9275
         X2              =   9275
         Y1              =   360
         Y2              =   2640
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         X1              =   11130
         X2              =   11130
         Y1              =   360
         Y2              =   2640
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   1855
         X2              =   1855
         Y1              =   360
         Y2              =   2640
      End
      Begin VB.Label lblJour 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   29
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblJour 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   2440
         TabIndex        =   31
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblJour 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   4295
         TabIndex        =   33
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblJour 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   6145
         TabIndex        =   36
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblJour 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   8000
         TabIndex        =   37
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblJour 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   9855
         TabIndex        =   40
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblJour 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   11810
         TabIndex        =   42
         Top             =   360
         Width           =   480
      End
   End
   Begin MSComCtl2.MonthView mvwSelection 
      Height          =   4020
      Left            =   120
      TabIndex        =   27
      Top             =   1200
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   7091
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   0
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowToday       =   0   'False
      StartOfWeek     =   90243073
      TitleBackColor  =   -2147483632
      CurrentDate     =   38353
   End
   Begin VB.Frame fraPunch 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   4800
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   8415
      Begin VB.ComboBox cmbType 
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   87
         Top             =   2520
         Width           =   5775
      End
      Begin VB.PictureBox picPMTypePunch 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   2640
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   79
         Top             =   840
         Width           =   1215
         Begin VB.OptionButton optTypePunch 
            Caption         =   "Électrique"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   81
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton optTypePunch 
            Caption         =   "Mécanique"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   80
            Top             =   300
            Width           =   1095
         End
      End
      Begin VB.OptionButton optHeureDiner 
         Caption         =   "1 heure"
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   5760
         TabIndex        =   73
         Top             =   4320
         Width           =   1215
      End
      Begin VB.OptionButton optHeureDiner 
         Caption         =   "30 minutes"
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   5760
         TabIndex        =   72
         Top             =   4080
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmdBrowseComment 
         Caption         =   "Choisir un commentaire"
         Height          =   375
         Left            =   2760
         TabIndex        =   71
         Top             =   4800
         Width           =   2055
      End
      Begin VB.CheckBox chkDiner 
         Caption         =   "Heure de dîner"
         Height          =   255
         Left            =   5040
         TabIndex        =   23
         Top             =   3720
         Width           =   1455
      End
      Begin VB.TextBox txtKM 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6360
         TabIndex        =   20
         Top             =   3360
         Width           =   615
      End
      Begin VB.CheckBox chkKM 
         Caption         =   "Kilométrage :"
         Height          =   255
         Left            =   5040
         TabIndex        =   19
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox txtClient 
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1860
         Width           =   5775
      End
      Begin VB.CommandButton cmdAnnuler 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Annuler"
         Height          =   375
         Left            =   6000
         TabIndex        =   25
         Top             =   4800
         Width           =   1095
      End
      Begin VB.TextBox txtCommentaires 
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   3720
         Width           =   4695
      End
      Begin VB.Frame fraHeure 
         Caption         =   "Heure"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   2055
         Begin MSMask.MaskEdBox mskHeure 
            Height          =   255
            Left            =   360
            TabIndex        =   15
            Top             =   600
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            _Version        =   393216
            PromptChar      =   "_"
         End
         Begin VB.OptionButton optHeure 
            Caption         =   "Système"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   16
            Top             =   600
            Width           =   255
         End
         Begin VB.OptionButton optHeure 
            Caption         =   "Heure de l'ordinateur"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.ComboBox cmbemployé 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1080
         TabIndex        =   8
         Text            =   "cmbemployé"
         Top             =   360
         Width           =   3495
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00C0C0C0&
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   7200
         TabIndex        =   26
         Top             =   4800
         Width           =   1095
      End
      Begin VB.TextBox txtEmploye 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   360
         Width           =   3495
      End
      Begin VB.TextBox txtNoProjet 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin MSMask.MaskEdBox mskNoProjet 
         Height          =   255
         Left            =   1200
         TabIndex        =   11
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   8
         Mask            =   "#####-##"
         PromptChar      =   "_"
      End
      Begin VB.Label lblTypePunch 
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4200
         TabIndex        =   89
         Top             =   840
         Width           =   3975
      End
      Begin VB.Label lblType 
         Caption         =   "Type"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2520
         TabIndex        =   86
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label lblPrefixe 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   960
         TabIndex        =   83
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "Km"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7080
         TabIndex        =   21
         Top             =   3360
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Commentaires"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label lblprojet 
         Caption         =   "No. Projet"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Employé"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Client"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2520
         TabIndex        =   17
         Top             =   1620
         Width           =   735
      End
   End
   Begin VB.Label lblNbreHeure 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   11400
      TabIndex        =   52
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblTitreHeure 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre d'heures dans la semaine pour :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5640
      TabIndex        =   50
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "frmPunch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Index de optHeure
Private Const I_OPT_SYSTEME              As Integer = 0
Private Const I_OPT_MANUELLEMENT         As Integer = 1

'Index de optTypePunch et optPMTypePunch
Private Const I_OPT_ELECTRIQUE           As Integer = 0
Private Const I_OPT_MECANIQUE            As Integer = 1

'Types quand c'est un 51
Private Const I_TYPE_ELEC_INSTALLATION   As Integer = 0
Private Const I_TYPE_ELEC_MISE_SERVICE   As Integer = 1

'Types quand c'est pas un 51
Private Const I_TYPE_ELEC_DESSIN         As Integer = 0
Private Const I_TYPE_ELEC_FABRICATION    As Integer = 1
Private Const I_TYPE_ELEC_ASSEMBLAGE     As Integer = 2
Private Const I_TYPE_ELEC_PROG_INTERFACE As Integer = 3
Private Const I_TYPE_ELEC_PROG_AUTOMATE  As Integer = 4
Private Const I_TYPE_ELEC_PROG_ROBOT     As Integer = 5
Private Const I_TYPE_ELEC_VISION         As Integer = 6
Private Const I_TYPE_ELEC_TEST           As Integer = 7
Private Const I_TYPE_ELEC_FORMATION      As Integer = 8
Private Const I_TYPE_ELEC_GESTION        As Integer = 9
Private Const I_TYPE_ELEC_SHIPPING       As Integer = 10
Private Const I_TYPE_ELEC_prototypage       As Integer = 11

'Types quand c'est un 51
Private Const I_TYPE_MEC_INSTALLATION    As Integer = 0

'Types quand c'est pas un 51
Private Const I_TYPE_MEC_DESSIN          As Integer = 0
Private Const I_TYPE_MEC_COUPE           As Integer = 1
Private Const I_TYPE_MEC_MACHINAGE       As Integer = 2
Private Const I_TYPE_MEC_SOUDURE         As Integer = 3
Private Const I_TYPE_MEC_ASSEMBLAGE      As Integer = 4
Private Const I_TYPE_MEC_PEINTURE        As Integer = 5
Private Const I_TYPE_MEC_TEST            As Integer = 6
Private Const I_TYPE_MEC_FORMATION       As Integer = 7
Private Const I_TYPE_MEC_GESTION         As Integer = 8
Private Const I_TYPE_MEC_SHIPPING        As Integer = 9
Private Const I_TYPE_MEC_prototypage        As Integer = 10

'Index de optHeureDiner
Private Const I_OPT_30_MINUTES           As Integer = 0
Private Const I_OPT_1_HEURE              As Integer = 1

'Index de lvwJour
Private Const I_LVW_NOM                  As Integer = 0
Private Const I_LVW_PROJET               As Integer = 1
Private Const I_LVW_DEBUT                As Integer = 2
Private Const I_LVW_FIN                  As Integer = 3
Private Const I_LVW_CLIENT               As Integer = 4
Private Const I_LVW_TYPE                 As Integer = 5
Private Const I_LVW_COMMENTAIRE          As Integer = 6
Private Const I_LVW_KM                   As Integer = 7

'Index de lvwJourSemaine
Private Const I_LVW_INITIALE             As Integer = 0
Private Const I_LVW_TEMPS                As Integer = 1

Private Enum enumPunch
  I_PUNCH_IN = 0
  I_PUNCH_OUT = 1
  I_MODIF_PUNCH_IN = 2
  I_MODIF_PUNCH_OUT = 3
End Enum

Private m_ePunch             As enumPunch
Private m_iNoEmploye         As Integer
Private m_datDateChoisie     As Date

Private m_bModifPunch        As Boolean
Private m_bMonthViewHasFocus As Boolean

Public sCommentaire          As String

Public Sub Afficher(ByVal sUserID As String)

5       On Error GoTo AfficherErreur

10      Dim rstEmploye As ADODB.Recordset
15      Dim iCompteur  As Integer

20      Call Unload(frmChoixPunch)

25      Set rstEmploye = New ADODB.Recordset

30      Call rstEmploye.Open("SELECT NoEmploye FROM GRB_Employés WHERE loginname = '" & sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)

35      m_iNoEmploye = rstEmploye.Fields("NoEmploye")

40      Call rstEmploye.Close
45      Set rstEmploye = Nothing

50      optHeure(I_OPT_SYSTEME).Value = True

55      mvwSelection.Year = Year(Date)
60      mvwSelection.Month = Month(Date)
65      mvwSelection.Day = Day(Date)

70      Call AfficherDate

75      Call RemplirComboEmploye

80      Call cmbHeureSemaine.Clear

85      For iCompteur = 0 To cmbemployé.ListCount - 1
90        Call cmbHeureSemaine.AddItem(cmbemployé.LIST(iCompteur))

95        cmbHeureSemaine.ItemData(cmbHeureSemaine.newIndex) = cmbemployé.ItemData(iCompteur)
100     Next

105     cmbHeureSemaine.ListIndex = 0

110     Call Me.Show

115     Exit Sub

AfficherErreur:

120     woups "frmPunch", "Afficher", Err, Erl
End Sub

Private Sub CalculerHeureSemaine()

5       On Error GoTo AfficherErreur

10      Dim rstPunch As ADODB.Recordset
15      Dim dblDebut As Double
20      Dim dblFin   As Double
25      Dim dblTotal As Double
30      Dim sDate    As String
35      Dim sDebut   As String
40      Dim sFin     As String

45      Set rstPunch = New ADODB.Recordset

50      Call rstPunch.Open("SELECT Date, HeureDébut, HeureFin FROM GRB_Punch WHERE NoEmploye = " & cmbHeureSemaine.ItemData(cmbHeureSemaine.ListIndex) & " AND Date BETWEEN '" & lvwJourSemaine(1).Tag & "' AND '" & lvwJourSemaine(7).Tag & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

55      Do While Not rstPunch.EOF
60        sDate = rstPunch.Fields("Date")

65        If Not IsNull(rstPunch.Fields("HeureDébut")) Then
70          If Trim(rstPunch.Fields("HeureDébut")) <> "" Then
75            sDebut = rstPunch.Fields("HeureDébut")
80          Else
85            sDebut = ""
90          End If
100       Else
105         sDebut = ""
110       End If

115       If Not IsNull(rstPunch.Fields("HeureFin")) Then
120         If Trim(rstPunch.Fields("HeureFin")) <> "" Then
125           sFin = rstPunch.Fields("HeureFin")
130         Else
135           sFin = ""
140         End If
145       Else
150         sFin = ""
155       End If

160       If sDebut <> "" And sFin <> "" Then
165         dblDebut = CDbl(Left$(sDebut, 2)) + CDbl(CDbl(Right$(sDebut, 2)) / 60)
170         dblFin = CDbl(Left$(sFin, 2)) + CDbl(CDbl(Right$(sFin, 2)) / 60)

175         dblTotal = dblTotal + (dblFin - dblDebut)
180       End If

185       Call rstPunch.MoveNext
190     Loop

195     Call rstPunch.Close
200     Set rstPunch = Nothing

205     lblNbreHeure.Caption = dblTotal

210     Exit Sub

AfficherErreur:

215     woups "frmPunch", "CalculerHeureSemaine", Err, Erl
End Sub

Private Sub AfficherDate()

5       On Error GoTo AfficherErreur

        'Affiche punch de la journée et de la semaine
        'dépendant la sélection dans le calendrier
10      Dim iCompteur As Integer

        'date choisie
15      m_datDateChoisie = DateSerial(mvwSelection.Year, mvwSelection.Month, mvwSelection.Day)

        'affiche punch jour et semaine
20      Call RemplirListViewJour
25      Call RemplirListViewJourAutorisation
30      Call RemplirListViewSemaine(False)
35      Call RemplirListViewSemaineAutorisation(False)

        'selectionne jour de la semaine
40      For iCompteur = 1 To 7
45        If lvwJourSemaine(iCompteur).Tag = m_datDateChoisie Then
50          lvwJourSemaine(iCompteur).BackColor = &HE0E0E0
55        Else
60          lvwJourSemaine(iCompteur).BackColor = &HFFFFFF
65        End If
70      Next

        'Affiche cedule une journee
75      frajour.Visible = True
80      fraPunch.Visible = False

85      Exit Sub

AfficherErreur:

90      woups "frmPunch", "AfficherDate", Err, Erl
End Sub

Private Sub RemplirListViewJour()

5       On Error GoTo AfficherErreur

        'remplis ListView une journée
10      Dim rstPunch   As ADODB.Recordset
15      Dim rstEmploye As ADODB.Recordset
20      Dim rstClient  As ADODB.Recordset
25      Dim itmPunch   As ListItem
30      Dim lForeColor As Long

        'vide le lister
35      Call lvwJour.ListItems.Clear

40      Set rstPunch = New ADODB.Recordset

45      Call rstPunch.Open("SELECT * FROM GRB_Punch WHERE Date = '" & ConvertDate(m_datDateChoisie) & "' AND NoEmploye = " & m_iNoEmploye & " ORDER BY HeureDébut", g_connData, adOpenDynamic, adLockOptimistic)

50      Set rstEmploye = New ADODB.Recordset
55      Set rstClient = New ADODB.Recordset

        'tant il y a de employé cedulé , ajoute dans lister
60      Do While Not rstPunch.EOF
65        Set itmPunch = lvwJour.ListItems.Add

70        Call rstEmploye.Open("SELECT initiale FROM GRB_Employés WHERE NoEmploye = " & m_iNoEmploye, g_connData, adOpenDynamic, adLockOptimistic)

75        itmPunch.Text = rstEmploye.Fields("initiale")

80        Call rstEmploye.Close

85        itmPunch.Tag = rstPunch.Fields("IDPunch")

90        If Not IsNull(rstPunch.Fields("NoProjet")) Then
95          itmPunch.SubItems(I_LVW_PROJET) = rstPunch.Fields("NoProjet")
100       Else
105         itmPunch.SubItems(I_LVW_PROJET) = vbNullString
110       End If

115       If Not IsNull(rstPunch.Fields("HeureDébut")) Then
120         itmPunch.SubItems(I_LVW_DEBUT) = rstPunch.Fields("HeureDébut")
125       Else
130         itmPunch.SubItems(I_LVW_DEBUT) = vbNullString
135       End If

140       If Not IsNull(rstPunch.Fields("HeureFin")) And rstPunch.Fields("HeureFin") <> vbNullString Then
145         itmPunch.SubItems(I_LVW_FIN) = rstPunch.Fields("HeureFin")
150         lForeColor = COLOR_NOIR
155       Else
160         itmPunch.SubItems(I_LVW_FIN) = vbNullString
165         lForeColor = COLOR_ROUGE
170       End If

175       If Not IsNull(rstPunch.Fields("NoClient")) And rstPunch.Fields("NoClient") <> vbNullString Then
180         Call rstClient.Open("SELECT NomClient FROM GRB_Client WHERE IDClient = " & rstPunch.Fields("NoClient"), g_connData, adOpenDynamic, adLockOptimistic)

185         itmPunch.SubItems(I_LVW_CLIENT) = rstClient.Fields("NomClient")

190         itmPunch.ListSubItems(I_LVW_CLIENT).Tag = rstPunch.Fields("NoClient")

195         Call rstClient.Close
200       Else
205         itmPunch.SubItems(I_LVW_CLIENT) = vbNullString
210       End If

215       If Not IsNull(rstPunch.Fields("Type")) Then
220         If Left$(itmPunch.SubItems(I_LVW_PROJET), 1) = "E" Then
225             itmPunch.SubItems(I_LVW_TYPE) = rstPunch.Fields("Type")
                
290
295         Else
300             itmPunch.SubItems(I_LVW_TYPE) = rstPunch.Fields("Type")
360         End If
365       End If

370       If Not IsNull(rstPunch.Fields("Commentaire")) And rstPunch.Fields("Commentaire") <> vbNullString Then
375         itmPunch.SubItems(I_LVW_COMMENTAIRE) = rstPunch.Fields("Commentaire")
380       Else
385         itmPunch.SubItems(I_LVW_COMMENTAIRE) = vbNullString
390       End If

395       If rstPunch.Fields("KM") = True Then
400         If Not IsNull(rstPunch.Fields("NbreKM")) Then
405           itmPunch.SubItems(I_LVW_KM) = rstPunch.Fields("NbreKM")
410         Else
415           itmPunch.SubItems(I_LVW_KM) = 0
420         End If
425       Else
430         itmPunch.SubItems(I_LVW_KM) = ""
435       End If

440       lvwJour.ListItems(itmPunch.Index).ForeColor = lForeColor
445       lvwJour.ListItems(itmPunch.Index).ListSubItems(I_LVW_PROJET).ForeColor = lForeColor
450       lvwJour.ListItems(itmPunch.Index).ListSubItems(I_LVW_DEBUT).ForeColor = lForeColor
455       lvwJour.ListItems(itmPunch.Index).ListSubItems(I_LVW_FIN).ForeColor = lForeColor
460       lvwJour.ListItems(itmPunch.Index).ListSubItems(I_LVW_CLIENT).ForeColor = lForeColor
465       lvwJour.ListItems(itmPunch.Index).ListSubItems(I_LVW_TYPE).ForeColor = lForeColor
470       lvwJour.ListItems(itmPunch.Index).ListSubItems(I_LVW_COMMENTAIRE).ForeColor = lForeColor
475       lvwJour.ListItems(itmPunch.Index).ListSubItems(I_LVW_KM).ForeColor = lForeColor

480       Call rstPunch.MoveNext
485     Loop

490     Set rstEmploye = Nothing
495     Set rstClient = Nothing

500     Call rstPunch.Close
505     Set rstPunch = Nothing

510     If lvwJour.ListItems.count > 0 Then
515       lvwJour.ListItems(lvwJour.ListItems.count).Selected = True
520     End If

525     Exit Sub

AfficherErreur:

530     woups "frmPunch", "RemplirListViewJour", Err, Erl
End Sub

Private Sub RemplirListViewJourAutorisation()

5       On Error GoTo AfficherErreur

        'Remplis ListView une journée
10      Dim rstPunch        As ADODB.Recordset
15      Dim rstEmploye      As ADODB.Recordset
20      Dim rstAutorisation As ADODB.Recordset
25      Dim rstClient       As ADODB.Recordset
30      Dim itmPunch        As ListItem
35      Dim lForeColor      As Long

40      Set rstAutorisation = New ADODB.Recordset

45      Call rstAutorisation.Open("SELECT * FROM GRB_AutorisationPunch WHERE AutoriserPar = " & m_iNoEmploye, g_connData, adOpenDynamic, adLockOptimistic)

50      Set rstPunch = New ADODB.Recordset
55      Set rstEmploye = New ADODB.Recordset
60      Set rstClient = New ADODB.Recordset

65      Do While Not rstAutorisation.EOF
70        Call rstPunch.Open("SELECT * FROM GRB_Punch WHERE Date = '" & ConvertDate(m_datDateChoisie) & "' AND NoEmploye = " & rstAutorisation.Fields("NoEmploye") & " ORDER BY HeureDébut", g_connData, adOpenDynamic, adLockOptimistic)

          'tant il y a de employé cedulé , ajoute dans lister
75        Do While Not rstPunch.EOF
80          Set itmPunch = lvwJour.ListItems.Add

85          Call rstEmploye.Open("SELECT initiale FROM GRB_Employés WHERE NoEmploye = " & rstAutorisation.Fields("NoEmploye"), g_connData, adOpenDynamic, adLockOptimistic)

90          itmPunch.Text = rstEmploye.Fields("initiale")

95          Call rstEmploye.Close

100         itmPunch.Tag = rstPunch.Fields("IDPunch")

105         If Not IsNull(rstPunch.Fields("NoProjet")) Then
110           itmPunch.SubItems(I_LVW_PROJET) = rstPunch.Fields("NoProjet")
115         Else
120           itmPunch.SubItems(I_LVW_PROJET) = vbNullString
125         End If

130         If Not IsNull(rstPunch.Fields("HeureDébut")) Then
135           itmPunch.SubItems(I_LVW_DEBUT) = rstPunch.Fields("HeureDébut")
140         Else
145           itmPunch.SubItems(I_LVW_DEBUT) = vbNullString
150         End If

155         If Not IsNull(rstPunch.Fields("HeureFin")) And rstPunch.Fields("HeureFin") <> vbNullString Then
160           itmPunch.SubItems(I_LVW_FIN) = rstPunch.Fields("HeureFin")
165           lForeColor = COLOR_NOIR
170         Else
175           itmPunch.SubItems(I_LVW_FIN) = vbNullString
180           lForeColor = COLOR_ROUGE
185         End If

190         If Not IsNull(rstPunch.Fields("NoClient")) And rstPunch.Fields("NoClient") <> vbNullString Then
195           Call rstClient.Open("SELECT NomClient FROM GRB_Client WHERE IDClient = " & rstPunch.Fields("NoClient"), g_connData, adOpenDynamic, adLockOptimistic)

200           itmPunch.SubItems(I_LVW_CLIENT) = rstClient.Fields("NomClient")

205           itmPunch.ListSubItems(I_LVW_CLIENT).Tag = rstPunch.Fields("NoClient")

210           Call rstClient.Close
215         Else
220           itmPunch.SubItems(I_LVW_CLIENT) = vbNullString
225         End If

230         If Not IsNull(rstPunch.Fields("Type")) Then
235           If Left$(itmPunch.SubItems(I_LVW_PROJET), 1) = "E" Then
240             Select Case rstPunch.Fields("Type")
                  Case "Dessin":        itmPunch.SubItems(I_LVW_TYPE) = "Dessins électriques"
245               Case "Fabrication":   itmPunch.SubItems(I_LVW_TYPE) = "Fabrication"
250               Case "Assemblage":    itmPunch.SubItems(I_LVW_TYPE) = "Assemblage du panneau"
255               Case "ProgInterface": itmPunch.SubItems(I_LVW_TYPE) = "Programmation d'interface"
260               Case "ProgAutomate":  itmPunch.SubItems(I_LVW_TYPE) = "Programmation d'automate"
265               Case "ProgRobot":     itmPunch.SubItems(I_LVW_TYPE) = "Programmation de robot"
270               Case "Vision":        itmPunch.SubItems(I_LVW_TYPE) = "Vision artificielle"
275               Case "Test":          itmPunch.SubItems(I_LVW_TYPE) = "Tests finaux"
280               Case "Installation":  itmPunch.SubItems(I_LVW_TYPE) = "Installation"
285               Case "MiseService":   itmPunch.SubItems(I_LVW_TYPE) = "Mise en service"
290               Case "Formation":     itmPunch.SubItems(I_LVW_TYPE) = "Formation du personnel"
295               Case "Gestion":       itmPunch.SubItems(I_LVW_TYPE) = "Gestion du projet"
300               Case "Shipping":      itmPunch.SubItems(I_LVW_TYPE) = "Expédition"
                  Case "Prototypage-Dévelloppement expérimental":      itmPunch.SubItems(I_LVW_TYPE) = "Prototypage-Dévelloppement expérimental"
305             End Select
310           Else
315             Select Case rstPunch.Fields("Type")
                  Case "Dessin":       itmPunch.SubItems(I_LVW_TYPE) = "Conception et dessins"
320               Case "Coupe":        itmPunch.SubItems(I_LVW_TYPE) = "Coupe et préparation (sauf soudage)"
325               Case "Machinage":    itmPunch.SubItems(I_LVW_TYPE) = "Machinage"
330               Case "Soudure":      itmPunch.SubItems(I_LVW_TYPE) = "Coupe, soudure et meulage"
335               Case "Assemblage":   itmPunch.SubItems(I_LVW_TYPE) = "Assemblage des systèmes"
340               Case "Peinture":     itmPunch.SubItems(I_LVW_TYPE) = "Peinture et finition"
345               Case "Test":         itmPunch.SubItems(I_LVW_TYPE) = "Tests finaux"
350               Case "Installation": itmPunch.SubItems(I_LVW_TYPE) = "Installation"
355               Case "Formation":    itmPunch.SubItems(I_LVW_TYPE) = "Formation du personnel"
360               Case "Gestion":      itmPunch.SubItems(I_LVW_TYPE) = "Gestion du projet"
365               Case "Shipping":     itmPunch.SubItems(I_LVW_TYPE) = "Expédition"
                  Case "Prototypage-Dévelloppement expérimental":      itmPunch.SubItems(I_LVW_TYPE) = "Prototypage-Dévelloppement expérimental"
370             End Select
375           End If
380         End If

385         If Not IsNull(rstPunch.Fields("Commentaire")) And rstPunch.Fields("Commentaire") <> vbNullString Then
390           itmPunch.SubItems(I_LVW_COMMENTAIRE) = rstPunch.Fields("Commentaire")
395         Else
400           itmPunch.SubItems(I_LVW_COMMENTAIRE) = vbNullString
405         End If

410         If rstPunch.Fields("KM") = True Then
415           itmPunch.SubItems(I_LVW_KM) = rstPunch.Fields("NbreKM")
420         Else
425           itmPunch.SubItems(I_LVW_KM) = vbNullString
430         End If

435         lvwJour.ListItems(itmPunch.Index).ForeColor = lForeColor
440         lvwJour.ListItems(itmPunch.Index).ListSubItems(I_LVW_PROJET).ForeColor = lForeColor
445         lvwJour.ListItems(itmPunch.Index).ListSubItems(I_LVW_DEBUT).ForeColor = lForeColor
450         lvwJour.ListItems(itmPunch.Index).ListSubItems(I_LVW_FIN).ForeColor = lForeColor
455         lvwJour.ListItems(itmPunch.Index).ListSubItems(I_LVW_CLIENT).ForeColor = lForeColor
460         lvwJour.ListItems(itmPunch.Index).ListSubItems(I_LVW_TYPE).ForeColor = lForeColor
465         lvwJour.ListItems(itmPunch.Index).ListSubItems(I_LVW_COMMENTAIRE).ForeColor = lForeColor
470         lvwJour.ListItems(itmPunch.Index).ListSubItems(I_LVW_KM).ForeColor = lForeColor

475         Call rstPunch.MoveNext
480       Loop

485       Call rstPunch.Close

490       Call rstAutorisation.MoveNext
495     Loop

500     Set rstPunch = Nothing
505     Set rstClient = Nothing
510     Set rstEmploye = Nothing

515     Call rstAutorisation.Close
520     Set rstAutorisation = Nothing

525     Call lvwJour_Click

530     Exit Sub

AfficherErreur:

535     woups "frmPunch", "RemplirListViewJourAutorisation", Err, Erl
End Sub

Private Sub RemplirListViewSemaine(ByVal bAujourdhui As Boolean)

5       On Error GoTo AfficherErreur

        'remplis une semaine
        'bAujourdhui sert à savoir si on rafraichit seulement la journée d'aujourd'hui
10      Dim rstPunch        As ADODB.Recordset
15      Dim rstEmploye      As ADODB.Recordset
20      Dim iJourSemaine    As Integer
25      Dim datPremiereDate As Date
30      Dim datDerniereDate As Date
35      Dim iCompteur       As Integer
40      Dim sHeureDebutFin  As String
45      Dim itmSemaine      As ListItem
50      Dim lForeColor      As Long

55      Set rstPunch = New ADODB.Recordset
60      Set rstEmploye = New ADODB.Recordset

65      If bAujourdhui = False Then
70        For iCompteur = 1 To 7
            'couleur par defaut entete de date
75          lbljour(iCompteur - 1).ForeColor = vbWhite
80          lblNomJour(iCompteur - 1).ForeColor = vbWhite

85          Call lvwJourSemaine(iCompteur).ListItems.Clear
90        Next

95        iJourSemaine = Weekday(m_datDateChoisie)
100       datPremiereDate = m_datDateChoisie
105       datDerniereDate = m_datDateChoisie

          'Trouve premiere date de la semaine
110       Do While Not Weekday(datPremiereDate) = 1
115         datPremiereDate = datPremiereDate - 1
120       Loop

          'Trouve derniere date de la semaine
125       Do While Not Weekday(datDerniereDate) = 7
130         datDerniereDate = datDerniereDate + 1
135       Loop

          'Sélectionne la semaine courante
140       Call rstPunch.Open("SELECT * FROM GRB_Punch WHERE Date BETWEEN '" & ConvertDate(datPremiereDate) & "' AND '" & ConvertDate(datDerniereDate) & "' AND NoEmploye = " & m_iNoEmploye & " ORDER BY HeureDébut", g_connData, adOpenDynamic, adLockOptimistic)

145       For iCompteur = 1 To 7
            'Pour écrire le jour
150         lbljour(iCompteur - 1).Caption = Day(datPremiereDate + iCompteur - 1)

            'Garde en memoire la date des listers
155         lvwJourSemaine(iCompteur).Tag = datPremiereDate + iCompteur - 1
160       Next

165       Do While Not rstPunch.EOF
            'ajoute dans le lister, dépendant le jour de la semaine
170         Set itmSemaine = lvwJourSemaine(Weekday(rstPunch.Fields("Date"))).ListItems.Add

175         itmSemaine.Tag = rstPunch.Fields("IDPunch")

180         Call rstEmploye.Open("SELECT initiale FROM GRB_Employés WHERE NoEmploye = " & m_iNoEmploye, g_connData, adOpenDynamic, adLockOptimistic)

185         itmSemaine.Text = rstEmploye.Fields("initiale")

190         Call rstEmploye.Close

195         sHeureDebutFin = Trim(rstPunch.Fields("HeureDébut") + "-")

200         If Not IsNull(rstPunch.Fields("HeureFin")) And rstPunch.Fields("HeureFin") <> vbNullString Then
205           sHeureDebutFin = sHeureDebutFin + rstPunch.Fields("HeureFin")
210           lForeColor = COLOR_NOIR
215         Else
220           lForeColor = COLOR_ROUGE
225         End If

230         itmSemaine.SubItems(I_LVW_TEMPS) = sHeureDebutFin

235         lvwJourSemaine(Weekday(rstPunch.Fields("Date"))).ListItems(itmSemaine.Index).ForeColor = lForeColor
240         lvwJourSemaine(Weekday(rstPunch.Fields("Date"))).ListItems(itmSemaine.Index).ListSubItems(I_LVW_TEMPS).ForeColor = lForeColor

245         Call rstPunch.MoveNext
250       Loop

255       Call rstPunch.Close
260     Else
265       Call lvwJourSemaine(Weekday(m_datDateChoisie)).ListItems.Clear

270       Call rstPunch.Open("SELECT * FROM GRB_Punch WHERE Date = '" & ConvertDate(m_datDateChoisie) & "' AND NoEmploye = " & m_iNoEmploye, g_connData, adOpenDynamic, adLockOptimistic)

275       Do While Not rstPunch.EOF
            'ajoute dans le lister, dépendant le jour de la semaine
280         Set itmSemaine = lvwJourSemaine(Weekday(rstPunch.Fields("Date"))).ListItems.Add

285         itmSemaine.Tag = rstPunch.Fields("IDPunch")

290         Call rstEmploye.Open("SELECT initiale FROM GRB_Employés WHERE NoEmploye = " & m_iNoEmploye, g_connData, adOpenDynamic, adLockOptimistic)

295         itmSemaine.Text = rstEmploye.Fields("initiale")

300         Call rstEmploye.Close

305         sHeureDebutFin = Trim(rstPunch.Fields("HeureDébut") + "-")

310         If Not IsNull(rstPunch.Fields("HeureFin")) And rstPunch.Fields("HeureFin") <> vbNullString Then
315           sHeureDebutFin = sHeureDebutFin + rstPunch.Fields("HeureFin")
320           lForeColor = COLOR_NOIR
325         Else
330           lForeColor = COLOR_ROUGE
335         End If

340         itmSemaine.SubItems(I_LVW_TEMPS) = sHeureDebutFin

345         lvwJourSemaine(Weekday(m_datDateChoisie)).ListItems(itmSemaine.Index).ForeColor = lForeColor
350         lvwJourSemaine(Weekday(m_datDateChoisie)).ListItems(itmSemaine.Index).ListSubItems(I_LVW_TEMPS).ForeColor = lForeColor

355         Call rstPunch.MoveNext
360       Loop

365       Call rstPunch.Close
370     End If

375     Set rstPunch = Nothing
380     Set rstEmploye = Nothing

385     Exit Sub

AfficherErreur:

390     woups "frmPunch", "RemplirListViewSemaine", Err, Erl
End Sub

Private Sub RemplirListViewSemaineAutorisation(ByVal bAujourdhui As Boolean)

5       On Error GoTo AfficherErreur

        'remplis une semaine
10      Dim rstPunch        As ADODB.Recordset
15      Dim rstEmploye      As ADODB.Recordset
20      Dim rstAutorisation As ADODB.Recordset
25      Dim iJourSemaine    As Integer
30      Dim datPremiereDate As Date
35      Dim datDerniereDate As Date
40      Dim iCompteur       As Integer
45      Dim sHeureDebutFin  As String
50      Dim itmSemaine      As ListItem
55      Dim lForeColor      As Long

60      Set rstPunch = New ADODB.Recordset
65      Set rstEmploye = New ADODB.Recordset
70      Set rstAutorisation = New ADODB.Recordset

75      If bAujourdhui = False Then
80        For iCompteur = 1 To 7
            'couleur par defaut entete de date
85          lbljour(iCompteur - 1).ForeColor = vbWhite
90          lblNomJour(iCompteur - 1).ForeColor = vbWhite
95        Next
    
100       iJourSemaine = Weekday(m_datDateChoisie)
105       datPremiereDate = m_datDateChoisie
110       datDerniereDate = m_datDateChoisie
    
          'trouve premiere date de la semaine
115       Do While Not Weekday(datPremiereDate) = 1
120         datPremiereDate = datPremiereDate - 1
125       Loop
    
          'trouve derniere date de la semaine
130       Do While Not Weekday(datDerniereDate) = 7
135         datDerniereDate = datDerniereDate + 1
140       Loop
    
145       For iCompteur = 1 To 7
            'pour ecrire le jour
150         lbljour(iCompteur - 1).Caption = Day(datPremiereDate + iCompteur - 1)
      
            'garde en memoire la date des lister
155         lvwJourSemaine(iCompteur).Tag = datPremiereDate + iCompteur - 1
160       Next
    
165       Call rstAutorisation.Open("SELECT * FROM GRB_AutorisationPunch WHERE AutoriserPar = " & m_iNoEmploye, g_connData, adOpenDynamic, adLockOptimistic)
     
170       Do While Not rstAutorisation.EOF
            'selectionne la semaine courante
175         Call rstPunch.Open("SELECT * FROM GRB_Punch WHERE Date BETWEEN '" & ConvertDate(datPremiereDate) & "' AND '" & ConvertDate(datDerniereDate) & "' AND NoEmploye = " & rstAutorisation.Fields("NoEmploye") & " ORDER BY HeureDébut", g_connData, adOpenDynamic, adLockOptimistic)
  
180         Do While Not rstPunch.EOF
              'ajoute dans le lister, dépendant le jour de la semaine
185           Set itmSemaine = lvwJourSemaine(Weekday(rstPunch.Fields("Date"))).ListItems.Add
      
190           itmSemaine.Tag = rstPunch.Fields("IDPunch")
       
195           Call rstEmploye.Open("SELECT initiale FROM GRB_Employés WHERE NoEmploye = " & rstAutorisation.Fields("NoEmploye"), g_connData, adOpenDynamic, adLockOptimistic)
       
200           itmSemaine.Text = rstEmploye.Fields("initiale")
        
205           Call rstEmploye.Close
               
210           sHeureDebutFin = Trim(rstPunch.Fields("HeureDébut") + "-")
              
215           If Not IsNull(rstPunch.Fields("HeureFin")) And rstPunch.Fields("HeureFin") <> vbNullString Then
220             sHeureDebutFin = sHeureDebutFin + rstPunch.Fields("HeureFin")
225             lForeColor = COLOR_NOIR
230           Else
235             lForeColor = COLOR_ROUGE
240           End If
      
245           itmSemaine.SubItems(I_LVW_TEMPS) = sHeureDebutFin
      
250           lvwJourSemaine(Weekday(rstPunch.Fields("Date"))).ListItems(itmSemaine.Index).ForeColor = lForeColor
255           lvwJourSemaine(Weekday(rstPunch.Fields("Date"))).ListItems(itmSemaine.Index).ListSubItems(I_LVW_TEMPS).ForeColor = lForeColor
      
260           Call rstPunch.MoveNext
265         Loop
        
270         Call rstPunch.Close
    
275         Call rstAutorisation.MoveNext
280       Loop
  
285       Call rstAutorisation.Close
295     Else
300       Call rstAutorisation.Open("SELECT * FROM GRB_AutorisationPunch WHERE AutoriserPar = " & m_iNoEmploye, g_connData, adOpenDynamic, adLockOptimistic)
     
305       Do While Not rstAutorisation.EOF
            'selectionne la semaine courante
310         Call rstPunch.Open("SELECT * FROM GRB_Punch WHERE Date = '" & ConvertDate(m_datDateChoisie) & "' AND NoEmploye = " & rstAutorisation.Fields("NoEmploye") & " ORDER BY HeureDébut", g_connData, adOpenDynamic, adLockOptimistic)
  
315         Do While Not rstPunch.EOF
              'ajoute dans le lister, dépendant le jour de la semaine
320           Set itmSemaine = lvwJourSemaine(Weekday(m_datDateChoisie)).ListItems.Add
      
325           itmSemaine.Tag = rstPunch.Fields("IDPunch")
       
330           Call rstEmploye.Open("SELECT initiale FROM GRB_Employés WHERE NoEmploye = " & rstAutorisation.Fields("NoEmploye"), g_connData, adOpenDynamic, adLockOptimistic)
       
335           itmSemaine.Text = rstEmploye.Fields("initiale")
        
340           Call rstEmploye.Close
               
345           sHeureDebutFin = Trim(rstPunch.Fields("HeureDébut") + "-")
              
350           If Not IsNull(rstPunch.Fields("HeureFin")) And rstPunch.Fields("HeureFin") <> vbNullString Then
355             sHeureDebutFin = sHeureDebutFin + rstPunch.Fields("HeureFin")
360             lForeColor = COLOR_NOIR
365           Else
370             lForeColor = COLOR_ROUGE
375           End If
      
380           itmSemaine.SubItems(I_LVW_TEMPS) = sHeureDebutFin
      
385           lvwJourSemaine(Weekday(m_datDateChoisie)).ListItems(itmSemaine.Index).ForeColor = lForeColor
390           lvwJourSemaine(Weekday(m_datDateChoisie)).ListItems(itmSemaine.Index).ListSubItems(I_LVW_TEMPS).ForeColor = lForeColor
      
395           Call rstPunch.MoveNext
400         Loop
        
405         Call rstPunch.Close
    
410         Call rstAutorisation.MoveNext
415       Loop
  
420       Call rstAutorisation.Close
425     End If

430     Set rstAutorisation = Nothing
435     Set rstEmploye = Nothing
440     Set rstPunch = Nothing

445     Exit Sub

AfficherErreur:

450     woups "frmPunch", "RemplirListViewSemaineAutorisation", Err, Erl
End Sub

Private Sub chkDiner_Click()

5       On Error GoTo AfficherErreur

10      If chkDiner.Value = vbChecked Then
15        optHeureDiner(I_OPT_1_HEURE).Enabled = True
20        optHeureDiner(I_OPT_30_MINUTES).Enabled = True
25      Else
30        optHeureDiner(I_OPT_1_HEURE).Enabled = False
35        optHeureDiner(I_OPT_30_MINUTES).Enabled = False
40      End If

45      m_bMonthViewHasFocus = False

50      Exit Sub

AfficherErreur:

55      woups "frmPunch", "chkDiner_Click", Err, Erl
End Sub

Private Sub chkPMDiner_Click()

5       On Error GoTo AfficherErreur

10      If chkPMDiner.Value = vbChecked Then
15        optPMHeureDiner(I_OPT_1_HEURE).Enabled = True
20        optPMHeureDiner(I_OPT_30_MINUTES).Enabled = True
25      Else
30        optPMHeureDiner(I_OPT_1_HEURE).Enabled = False
35        optPMHeureDiner(I_OPT_30_MINUTES).Enabled = False
40      End If

50      m_bMonthViewHasFocus = False

55      Exit Sub

AfficherErreur:

60      woups "frmPunch", "chkPMDiner_Click", Err, Erl
End Sub

Private Sub chkDiner_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5       On Error GoTo AfficherErreur

10      If m_bMonthViewHasFocus = True Then
15        If chkDiner.Value = vbChecked Then
20          chkDiner.Value = vbUnchecked
25        Else
30          chkDiner.Value = vbChecked
35        End If
40      End If

45      Exit Sub

AfficherErreur:

50      woups "frmPunch", "chkDiner_MouseUp", Err, Erl
End Sub

Private Sub chkPMDiner_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5       On Error GoTo AfficherErreur

10      If m_bMonthViewHasFocus = True Then
15        If chkPMDiner.Value = vbChecked Then
20          chkPMDiner.Value = vbUnchecked
25        Else
30          chkPMDiner.Value = vbChecked
35        End If
40      End If

45      Exit Sub

AfficherErreur:

50      woups "frmPunch", "chkPMDiner_MouseUp", Err, Erl
End Sub

Private Sub chkKM_Click()

5       On Error GoTo AfficherErreur

10      If chkKM.Value = vbChecked Then
15        txtKM.Enabled = True
20      Else
25        txtKM.Text = ""
30        txtKM.Enabled = False
35      End If

40      m_bMonthViewHasFocus = False

45      Exit Sub

AfficherErreur:

50      woups "frmPunch", "chkKM_Click", Err, Erl
End Sub

Private Sub chkKM_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5       On Error GoTo AfficherErreur

10      If m_bMonthViewHasFocus = True Then
15        If chkKM.Value = vbChecked Then
20          chkKM.Value = vbUnchecked
25        Else
30          chkKM.Value = vbChecked
35        End If
40      End If

45      Exit Sub

AfficherErreur:

50      woups "frmPunch", "chkKM_MouseUp", Err, Erl
End Sub

Private Sub cmbEmployé_Click()

5       On Error GoTo AfficherErreur

10      txtEmploye.Text = cmbemployé.Text

15      Exit Sub

AfficherErreur:

20      woups "frmPunch", "cmbEmployé_Click", Err, Erl
End Sub

Private Sub cmbHeureSemaine_Click()
        
5       On Error GoTo AfficherErreur

10      Call CalculerHeureSemaine

15      Exit Sub

AfficherErreur:

20      woups "frmPunch", "cmbHeureSemaine_Click", Err, Erl
End Sub

Private Sub cmdAnnuler_Click()

5       On Error GoTo AfficherErreur

10      frajour.Visible = True
15      fraPunch.Visible = False

20      m_bMonthViewHasFocus = False

25      Exit Sub

AfficherErreur:

30      woups "frmPunch", "cmdAnnuler_Click", Err, Erl
End Sub

Private Sub cmdBrowseComment_Click()

5       On Error GoTo AfficherErreur

10      Dim sProjet As String

15      If mskNoProjet.Visible = True Then
20        sProjet = mskNoProjet.Text
25      Else
30        sProjet = txtnoprojet.Text
35      End If

40      If txtnoprojet.Text <> "" Or mskNoProjet.Text <> "" Then
45        If txtClient.Text <> "" Then
50          Call frmChoixCommentaire.Afficher(sProjet)

55          If sCommentaire <> "" Then
60            txtCommentaires.Text = sCommentaire
65          End If
70        Else
75          Call MsgBox("Numéro de projet ou soumission invalide!", vbOKOnly, "Erreur")
80        End If
85      Else
90        Call MsgBox("Le numéro de projet ou soumission est obligatoire!", vbOKOnly, "Erreur")
95      End If

100     Exit Sub

AfficherErreur:

105     woups "frmPunch", "cmdBrowseComment_Click", Err, Erl
End Sub

Private Sub cmdBrowseCommentPM_Click()

5       On Error GoTo AfficherErreur

10      If mskPMNoProjet.Text <> "" Then
15        If txtPMClient.Text <> "" Then
20          Call frmChoixCommentaire.Afficher(mskPMNoProjet.Text)

25          txtCommentaires.Text = sCommentaire
30        Else
35          Call MsgBox("Numéro de projet ou soumission invalide!", vbOKOnly, "Erreur")
40        End If
45      Else
50        Call MsgBox("Le numéro de projet ou soumission est obligaoire!", vbOKOnly, "Erreur")
55      End If

60      Exit Sub

AfficherErreur:

65      woups "frmPunch", "cmdBrowseCommentPM_Click", Err, Erl
End Sub

Private Sub cmdPMAnnuler_Click()

5       On Error GoTo AfficherErreur

10      frajour.Visible = True
15      fraPunch.Visible = False
20      fraPunchMultiple.Visible = False

25      m_bMonthViewHasFocus = False

30      Exit Sub

AfficherErreur:

35      woups "frmPunch", "cmdPMAnnuler_Click", Err, Erl
End Sub

Private Sub cmdAnnuler_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5       On Error GoTo AfficherErreur

10      If m_bMonthViewHasFocus = True Then
15        Call cmdAnnuler_Click
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmPunch", "cmdAnnuler_MouseUp", Err, Erl
End Sub

Private Sub cmdPMAnnuler_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5       On Error GoTo AfficherErreur

10      If m_bMonthViewHasFocus = True Then
15        Call cmdPMAnnuler_Click
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmPunch", "cmdPMAnnuler_MouseUp", Err, Erl
End Sub

Private Sub cmdModifierPunchIn_Click()

5       On Error GoTo AfficherErreur

10      Dim rstEmploye As ADODB.Recordset
15      Dim itmPunch   As ListItem
20      Dim iCompteur  As Integer

25      If VerifierModificationDate = True Then
30        Set itmPunch = lvwJour.SelectedItem
  
35        Set rstEmploye = New ADODB.Recordset
    
40        Call rstEmploye.Open("SELECT employe FROM GRB_Employés WHERE Initiale = '" & itmPunch.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
          
45        For iCompteur = 0 To cmbemployé.ListCount - 1
50          If cmbemployé.LIST(iCompteur) = rstEmploye.Fields("employe") Then
55            cmbemployé.ListIndex = iCompteur
              
60            Exit For
65          End If
70        Next
   
75        Call rstEmploye.Close
80        Set rstEmploye = Nothing
    
85        txtClient.Text = itmPunch.SubItems(I_LVW_CLIENT)
  
90        txtClient.Tag = itmPunch.ListSubItems(I_LVW_CLIENT).Tag
    
95        cmbemployé.Visible = True
100       txtEmploye.Visible = False
            
105       Select Case Left$(itmPunch.SubItems(I_LVW_PROJET), 1)
            Case "E": optTypePunch(I_OPT_ELECTRIQUE).Value = True
110         Case "M": optTypePunch(I_OPT_MECANIQUE).Value = True
115       End Select
  
120       mskNoProjet.Text = Right$(itmPunch.SubItems(I_LVW_PROJET), 8)
            
125       m_ePunch = I_MODIF_PUNCH_IN
            
130       Call RemplirComboType
            
135       If Left$(itmPunch.SubItems(I_LVW_PROJET), 1) = "E" Then
            
140         If Not IsNull(itmPunch.SubItems(I_LVW_TYPE)) Then
                cmbType.Text = itmPunch.SubItems(I_LVW_TYPE)
            Else
            
                cmbType.ListIndex = -1
210         End If

215       Else
        If Not itmPunch.SubItems(I_LVW_TYPE) = "Soumission" Then
            If Not IsNull(itmPunch.SubItems(I_LVW_TYPE)) Then
                cmbType.Text = itmPunch.SubItems(I_LVW_TYPE)
                
            Else
                cmbType.ListIndex = -1
285       End If
        End If
    End If
            
290       mskNoProjet.Visible = True
295       txtnoprojet.Visible = False
    
300       picTypePunch.Enabled = True
  
305       mskHeure.mask = "##:##"
310       mskHeure.Text = itmPunch.SubItems(I_LVW_DEBUT)
  
315       m_bModifPunch = True
  
320       optHeure(I_OPT_MANUELLEMENT).Value = True
  
325       m_bModifPunch = False
      
330       txtCommentaires.Text = itmPunch.SubItems(I_LVW_COMMENTAIRE)
  
335       If itmPunch.SubItems(I_LVW_KM) <> "" Then
340         chkKM.Value = vbChecked
345         txtKM.Text = itmPunch.SubItems(I_LVW_KM)
350       Else
355         chkKM.Value = vbUnchecked
360         txtKM.Text = vbNullString
365       End If
    
370       fraPunch.Caption = "Modification du punch in"
    
375       frajour.Visible = False
380       fraPunchMultiple.Visible = False
385       fraPunch.Visible = True
  
390       chkDiner.Visible = False
395       optHeureDiner(I_OPT_30_MINUTES).Visible = False
400       optHeureDiner(I_OPT_1_HEURE).Visible = False
405     End If

410     m_bMonthViewHasFocus = False

415     Exit Sub

AfficherErreur:

420     woups "frmPunch", "cmdModifierPunchIn_Click", Err, Erl
End Sub

Private Sub cmdModifierPunchIn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5       On Error GoTo AfficherErreur

10      If m_bMonthViewHasFocus = True Then
15        Call cmdModifierPunchIn_Click
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmPunch", "cmdModifierPunchIn_MouseUp", Err, Erl
End Sub

Private Sub cmdModifierPunchOut_Click()

5       On Error GoTo AfficherErreur

10      Dim rstEmploye As ADODB.Recordset

15      If VerifierModificationDate = True Then
20        If lvwJour.ListItems.count > 0 Then
25          m_ePunch = I_MODIF_PUNCH_OUT
      
30          Call AfficherPunchOut
  
35          fraPunch.Caption = "Modification du punch out"
  
40          chkDiner.Visible = True
45          optHeureDiner(I_OPT_30_MINUTES).Visible = True
50          optHeureDiner(I_OPT_1_HEURE).Visible = True
  
55          Set rstEmploye = New ADODB.Recordset
  
60          Call rstEmploye.Open("SELECT GRB_Famille.Famille FROM GRB_employés INNER JOIN GRB_Famille ON GRB_employés.Famille = GRB_Famille.IDFamille WHERE employe = '" & m_iNoEmploye & "'", g_connData, adOpenDynamic, adLockOptimistic)
    
65          If Not rstEmploye.EOF Then
70            Select Case rstEmploye.Fields("Famille")
                Case "Administration": optHeureDiner(I_OPT_1_HEURE).Value = True
75              Case "Technicien":     optHeureDiner(I_OPT_1_HEURE).Value = True
80              Case Else:             optHeureDiner(I_OPT_30_MINUTES).Value = True
85            End Select
90          Else
95            optHeureDiner(I_OPT_30_MINUTES).Value = True
100         End If
  
105         Call rstEmploye.Close
110         Set rstEmploye = Nothing
  
115         chkDiner.Value = vbUnchecked
  
120         m_bModifPunch = True
    
125         optHeure(I_OPT_MANUELLEMENT).Value = True
  
130         m_bModifPunch = False
        
135         mskHeure.mask = "##:##"
140         mskHeure.Text = lvwJour.SelectedItem.SubItems(I_LVW_FIN)
145       End If
150     End If

155     m_bMonthViewHasFocus = False

160     Exit Sub

AfficherErreur:

165     woups "frmPunch", "cmdModifierPunchOut_Click", Err, Erl
End Sub

Private Sub cmdModifierPunchOut_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5       On Error GoTo AfficherErreur

10      If m_bMonthViewHasFocus = True Then
15        Call cmdModifierPunchOut_Click
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmPunch", "cmdModifierPunchOut_MouseUp", Err, Erl
End Sub

Private Sub cmdOK_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
5       On Error GoTo AfficherErreur

10      If m_bMonthViewHasFocus = True Then
15        Call cmdOK_Click
20      End If

25      Exit Sub
  
AfficherErreur:

30      woups "frmPunch", "cmdOK_MouseUp", Err, Erl
End Sub

Private Sub cmdPMOK_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
5       On Error GoTo AfficherErreur

10      If m_bMonthViewHasFocus = True Then
15        Call cmdPMOK_Click
20      End If

25      Exit Sub
  
AfficherErreur:

30      woups "frmPunch", "cmdPMOK_MouseUp", Err, Erl
End Sub

Private Sub cmdPunchIn_Click()

5       On Error GoTo AfficherErreur
  
10      Dim rstEmploye As ADODB.Recordset
15      Dim iCompteur  As Integer
  
20      If VerifierModificationDate = True Then
25        mskNoProjet.mask = vbNullString
30        mskNoProjet.Text = vbNullString
35        mskNoProjet.mask = "#####-##"
    
40        txtClient.Text = vbNullString
    
45        cmbemployé.Visible = True
50        txtEmploye.Visible = False
    
55        mskNoProjet.Visible = True
60        txtnoprojet.Visible = False
    
65        picTypePunch.Enabled = True
  
70        Set rstEmploye = New ADODB.Recordset
  
75        Call rstEmploye.Open("SELECT GRB_Famille.Famille FROM GRB_employés INNER JOIN GRB_Famille ON GRB_employés.Famille = GRB_Famille.IDFamille WHERE employe = '" & m_iNoEmploye & "'", g_connData, adOpenDynamic, adLockOptimistic)
    
80        If Not rstEmploye.EOF Then
85          Select Case rstEmploye.Fields("Famille")
              Case "Électrique": optTypePunch(I_OPT_ELECTRIQUE).Value = True
90            Case "Mécanique":  optTypePunch(I_OPT_MECANIQUE).Value = True
95            Case Else:         optTypePunch(I_OPT_ELECTRIQUE).Value = True
100         End Select
105       Else
110         optTypePunch(I_OPT_ELECTRIQUE).Value = True
115       End If
  
120       Call rstEmploye.Close
125       Set rstEmploye = Nothing
  
130       cmbType.ListIndex = -1
  
    
135       mskHeure.mask = vbNullString
140       mskHeure.Text = vbNullString
145       mskHeure.mask = "##:##"
    
150       optHeure(I_OPT_SYSTEME).Value = True
    
155       txtCommentaires.Text = vbNullString
  
160       chkKM.Value = vbUnchecked
  
165       txtKM.Text = vbNullString
    
170       fraPunch.Caption = "Punch in"
    
175       frajour.Visible = False
180       fraPunch.Visible = True
185       fraPunchMultiple.Visible = False
  
190       chkDiner.Visible = False
195       optHeureDiner(I_OPT_30_MINUTES).Visible = False
200       optHeureDiner(I_OPT_1_HEURE).Visible = False
  
205       Call mskNoProjet.SetFocus
  
210       m_ePunch = I_PUNCH_IN
215     End If

220     m_bMonthViewHasFocus = False

225     Exit Sub

AfficherErreur:

230     woups "frmPunch", "cmdPunchIn_Click", Err, Erl
End Sub

Private Sub cmdPunchIn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5       On Error GoTo AfficherErreur

10      If m_bMonthViewHasFocus = True Then
15        Call cmdPunchIn_Click
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmPunch", "cmdPunchIn_MouseUp", Err, Erl
End Sub

Private Sub cmdPunchMultiple_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5       On Error GoTo AfficherErreur

10      If m_bMonthViewHasFocus = True Then
15        Call cmdPunchMultiple_Click
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmPunch", "cmdPunchMultiple_MouseUp", Err, Erl
End Sub

Private Sub cmdPunchMultiple_Click()

5       On Error GoTo AfficherErreur

10      Dim iCompteur  As Integer
15      Dim rstEmploye As ADODB.Recordset

20      If VerifierModificationDate = True Then
25        For iCompteur = 1 To lvwEmployes.ListItems.count
30          lvwEmployes.ListItems(iCompteur).Checked = False
35        Next
  
40        mskPMNoProjet.mask = vbNullString
45        mskPMNoProjet.Text = vbNullString
50        mskPMNoProjet.mask = "#####-##"
    
55        Set rstEmploye = New ADODB.Recordset

60        Call rstEmploye.Open("SELECT GRB_Famille.Famille FROM GRB_employés INNER JOIN GRB_Famille ON GRB_employés.Famille = GRB_Famille.IDFamille WHERE employe = '" & m_iNoEmploye & "'", g_connData, adOpenDynamic, adLockOptimistic)

65        If Not rstEmploye.EOF Then
70          Select Case rstEmploye.Fields("Famille")
              Case "Électrique": optPMTypePunch(I_OPT_ELECTRIQUE).Value = True
75            Case "Mécanique":  optPMTypePunch(I_OPT_MECANIQUE).Value = True
80            Case Else:         optPMTypePunch(I_OPT_ELECTRIQUE).Value = True
85          End Select
90        Else
95          optPMTypePunch(I_OPT_ELECTRIQUE).Value = True
100       End If

105       Call rstEmploye.Close
110       Set rstEmploye = Nothing
  
115       mskPMHeureDebut.mask = vbNullString
120       mskPMHeureDebut.Text = vbNullString
125       mskPMHeureDebut.mask = "##:##"

130       mskPMHeureFin.mask = vbNullString
135       mskPMHeureFin.Text = vbNullString
140       mskPMHeureFin.mask = "##:##"
    
145       txtPMCommentaire.Text = vbNullString
  
150       chkPMDiner.Value = vbUnchecked

155       fraPunch.Visible = False
160       frajour.Visible = False
165       fraPunchMultiple.Visible = True
170     End If

175     m_bMonthViewHasFocus = False

180     Exit Sub

AfficherErreur:

185     woups "frmPunch", "cmdPunchMultiple_Click", Err, Erl
End Sub

Private Sub cmdPunchOut_Click()

5       On Error GoTo AfficherErreur

10      Dim rstEmploye As ADODB.Recordset

15      If VerifierModificationDate = True Then
20        If lvwJour.ListItems.count > 0 Then
25          If lvwJour.SelectedItem.ForeColor = COLOR_ROUGE Then
30            m_ePunch = I_PUNCH_OUT
      
35            Call AfficherPunchOut
  
40            fraPunch.Caption = "Punch out"
  
45            chkDiner.Visible = True
50            optHeureDiner(I_OPT_30_MINUTES).Visible = True
55            optHeureDiner(I_OPT_1_HEURE).Visible = True
  
60            Set rstEmploye = New ADODB.Recordset
  
65            Call rstEmploye.Open("SELECT GRB_Famille.Famille FROM GRB_employés INNER JOIN GRB_Famille ON GRB_employés.Famille = GRB_Famille.IDFamille WHERE employe = '" & m_iNoEmploye & "'", g_connData, adOpenDynamic, adLockOptimistic)
    
70            If Not rstEmploye.EOF Then
75              Select Case rstEmploye.Fields("Famille")
                  Case "Administration": optHeureDiner(I_OPT_1_HEURE).Value = True
80                Case "Technicien":     optHeureDiner(I_OPT_1_HEURE).Value = True
85                Case Else:             optHeureDiner(I_OPT_30_MINUTES).Value = True
90              End Select
95            Else
100             optHeureDiner(I_OPT_30_MINUTES).Value = True
105           End If
  
110           Call rstEmploye.Close
115           Set rstEmploye = Nothing
  
120           chkDiner.Value = vbUnchecked
       
125           mskHeure.mask = vbNullString
130           mskHeure.Text = vbNullString
135           mskHeure.mask = "##:##"
        
140           optHeure(I_OPT_SYSTEME).Value = True
145         Else
150           Call MsgBox("Le punch out a déjà été fait!")
155         End If
160       End If
165     End If

170     m_bMonthViewHasFocus = False

175     Exit Sub

AfficherErreur:

180     woups "frmPunch", "cmdPunchOut_Click", Err, Erl
End Sub

Private Sub AfficherPunchOut()

5       On Error GoTo AfficherErreur

10      Dim rstEmploye As ADODB.Recordset
15      Dim rstPunch   As ADODB.Recordset
20      Dim rstClient  As ADODB.Recordset
25      Dim iCompteur  As Integer
        Dim G As Integer
30      Set rstPunch = New ADODB.Recordset
35      Set rstEmploye = New ADODB.Recordset
40      Set rstClient = New ADODB.Recordset
    
45      Call rstPunch.Open("SELECT * FROM GRB_Punch WHERE IDPunch = " & lvwJour.SelectedItem.Tag, g_connData, adOpenDynamic, adLockOptimistic)
  
50      Call rstEmploye.Open("SELECT employe FROM GRB_Employés WHERE NoEmploye = " & rstPunch.Fields("NoEmploye"), g_connData, adOpenDynamic, adLockOptimistic)
  
55      For iCompteur = 0 To cmbemployé.ListCount - 1
60        If cmbemployé.LIST(iCompteur) = rstEmploye.Fields("Employe") Then
65          cmbemployé.ListIndex = iCompteur

70          Exit For
75        End If
80      Next
  
85      Call rstEmploye.Close
90      Set rstEmploye = Nothing
    
95      Call rstClient.Open("SELECT NomClient FROM GRB_Client WHERE IDClient = " & rstPunch.Fields("NoClient"), g_connData, adOpenDynamic, adLockOptimistic)
    
100     txtClient.Text = rstClient.Fields("NomClient")

105     txtClient.Tag = rstPunch.Fields("NoClient")
  
110     Call rstClient.Close
115     Set rstClient = Nothing
        
120     txtnoprojet.Text = Right(rstPunch.Fields("NoProjet"), 8)
  
125     Call RemplirComboType
  
130     Call AfficherTypePunch
  
135     If Not IsNull(rstPunch.Fields("Commentaire")) Then
140       txtCommentaires.Text = rstPunch.Fields("Commentaire")
145     Else
150       txtCommentaires.Text = vbNullString
155     End If

160     If rstPunch.Fields("KM") = True Then
165       chkKM.Value = vbChecked

170       If Not IsNull(rstPunch.Fields("NbreKM")) Then
175         txtKM.Text = rstPunch.Fields("NbreKM")
180       Else
185         txtKM.Text = 0
190       End If
195     Else
200       chkKM.Value = vbUnchecked
205       txtKM.Text = vbNullString
210     End If
   
215     Select Case Left(rstPunch.Fields("NoProjet"), 1)
          Case "E": optTypePunch(I_OPT_ELECTRIQUE).Value = True
220       Case "M": optTypePunch(I_OPT_MECANIQUE).Value = True
225     End Select

230     If Not IsNull(rstPunch.Fields("Type")) Then
235       If Left(rstPunch.Fields("NoProjet"), 1) = "E" Then
                     For G = 0 To cmbType.ListCount
                        If cmbType.LIST(G) = rstPunch.Fields("Type") Then
                            cmbType.ListIndex = G
                            Exit For
                        End If
                    Next
310
315       Else
                    For G = 0 To cmbType.ListCount
                        If cmbType.LIST(G) = rstPunch.Fields("Type") Then
                            cmbType.ListIndex = G
                            Exit For
                        End If
                    Next
385       End If
390     End If
            
395     picTypePunch.Enabled = False

400     txtnoprojet.Visible = True
405     mskNoProjet.Visible = False
  
410     txtEmploye.Visible = True
415     cmbemployé.Visible = False
  
420     frajour.Visible = False
425     fraPunch.Visible = True
430     fraPunchMultiple.Visible = False

435     Exit Sub

AfficherErreur:

440     woups "frmPunch", "AfficherPunchOut", Err, Erl
End Sub

Private Sub RemplirComboEmploye()

5       On Error GoTo AfficherErreur

10      Dim rstEmploye      As ADODB.Recordset
15      Dim rstAutorisation As ADODB.Recordset
20      Dim itmEmploye      As ListItem
  
25      Call cmbemployé.Clear
30      Call lvwEmployes.ListItems.Clear
  
35      Set rstEmploye = New ADODB.Recordset
40      Set rstAutorisation = New ADODB.Recordset
  
45      Call rstEmploye.Open("SELECT Employe FROM GRB_Employés WHERE NoEmploye = " & m_iNoEmploye, g_connData, adOpenDynamic, adLockOptimistic)
  
50      Call cmbemployé.AddItem(rstEmploye.Fields("Employe"))
  
55      cmbemployé.ItemData(cmbemployé.newIndex) = m_iNoEmploye

60      Set itmEmploye = lvwEmployes.ListItems.Add

65      itmEmploye.Text = rstEmploye.Fields("Employe")
70      itmEmploye.Tag = m_iNoEmploye
  
75      Call rstEmploye.Close
  
80      Call rstAutorisation.Open("SELECT * FROM GRB_AutorisationPunch WHERE AutoriserPar = " & m_iNoEmploye, g_connData, adOpenDynamic, adLockOptimistic)
  
85      If Not rstAutorisation.EOF Then
90        Do While Not rstAutorisation.EOF
95          Call rstEmploye.Open("SELECT Employe FROM GRB_Employés WHERE NoEmploye = " & rstAutorisation.Fields("NoEmploye"), g_connData, adOpenDynamic, adLockOptimistic)
    
100         Call cmbemployé.AddItem(rstEmploye.Fields("Employe"))
    
105         cmbemployé.ItemData(cmbemployé.newIndex) = rstAutorisation.Fields("NoEmploye")

110         Set itmEmploye = lvwEmployes.ListItems.Add

115         itmEmploye.Text = rstEmploye.Fields("Employe")
120         itmEmploye.Tag = rstAutorisation.Fields("NoEmploye")
    
125         Call rstEmploye.Close
    
130         Call rstAutorisation.MoveNext
135       Loop

140       cmdPunchMultiple.Visible = True
145     Else
150       cmdPunchMultiple.Visible = False 'Gll
155     End If

160     Call rstAutorisation.Close
165     Set rstAutorisation = Nothing

170     Set rstEmploye = Nothing

175     If cmbemployé.ListCount = 1 Then
180       cmbemployé.ListIndex = 0
185     End If

190     Exit Sub

AfficherErreur:

195     woups "frmPunch", "RemplirComboEmploye", Err, Erl
End Sub

Private Sub cmdPunchOut_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5       On Error GoTo AfficherErreur

10      If m_bMonthViewHasFocus = True Then
15        Call cmdPunchOut_Click
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmPunch", "cmdPunchOut_MouseUp", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur
        Call Table_exist
10      mvwSelection.StartOfWeek = vbSunday
    
15      Exit Sub

AfficherErreur:

20      woups "frmPunch", "Form_Load", Err, Erl
End Sub
Private Sub lvwJour_Click()

5       On Error GoTo AfficherErreur

10      If lvwJour.ListItems.count > 0 Then
15        If lvwJour.SelectedItem.ForeColor = COLOR_ROUGE Then
20          cmdModifierPunchIn.Enabled = True
25          cmdModifierPunchOut.Enabled = False
30        Else
35          cmdModifierPunchIn.Enabled = True
40          cmdModifierPunchOut.Enabled = True
45        End If
50      Else
55        cmdModifierPunchIn.Enabled = False
60        cmdModifierPunchOut.Enabled = False
65      End If

70      Exit Sub

AfficherErreur:

75      woups "frmPunch", "lvwJour_Click", Err, Erl
End Sub

Private Sub lvwJour_KeyDown(KeyCode As Integer, Shift As Integer)

5       On Error GoTo AfficherErreur

10      If lvwJour.ListItems.count > 0 Then
15        If KeyCode = vbKeyDelete Then
20          If MsgBox("Voulez-vous vraiment effacer ce punch ?", vbYesNo) = vbYes Then
25            Call EffacerPunch
30          End If
35        End If
40      End If

45      Exit Sub

AfficherErreur:

50      woups "frmPunch", "lvwJour_KeyDown", Err, Erl
End Sub

Private Sub EffacerPunch()

5       On Error GoTo AfficherErreur

        'Efface le punch sélectionné
10      Call g_connData.Execute("DELETE * FROM GRB_Punch WHERE IDPunch = " & lvwJour.SelectedItem.Tag)
  
15      Call RemplirListViewSemaine(False)
20      Call RemplirListViewSemaineAutorisation(False)
25      Call RemplirListViewJour
30      Call RemplirListViewJourAutorisation

35      Call CalculerHeureSemaine

40      Exit Sub

AfficherErreur:

45      woups "frmPunch", "EffacerPunch", Err, Erl
End Sub

Private Sub mskHeure_GotFocus()

5       On Error GoTo AfficherErreur

10      mskHeure.mask = "##:##"

15      Exit Sub

AfficherErreur:

20      woups "frmPunch", "mskHeure_GotFocus", Err, Erl
End Sub

Private Sub mskHeure_LostFocus()

5       On Error GoTo AfficherErreur

10      mskHeure.mask = vbNullString
  
15      If mskHeure.Text = "__:__" Then
20        mskHeure.Text = vbNullString
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmPunch", "mskHeure_LostFocus", Err, Erl
End Sub

Private Sub mskPMHeureDebut_GotFocus()

5       On Error GoTo AfficherErreur

10      mskPMHeureDebut.mask = "##:##"

15      Exit Sub

AfficherErreur:

20      woups "frmPunch", "mskPMHeureDebut_GotFocus", Err, Erl
End Sub

Private Sub mskPMHeureDebut_LostFocus()

5       On Error GoTo AfficherErreur

10      mskPMHeureDebut.mask = vbNullString

15      If mskPMHeureDebut.Text = "__:__" Then
20        mskPMHeureDebut.Text = vbNullString
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmPunch", "mskPMHeureDebut_LostFocus", Err, Erl
End Sub

Private Sub mskPMHeureFin_GotFocus()

5       On Error GoTo AfficherErreur

10      mskPMHeureFin.mask = "##:##"

15      Exit Sub

AfficherErreur:

20      woups "frmPunch", "mskPMHeureFin_GotFocus", Err, Erl
End Sub

Private Sub mskPMHeureFin_LostFocus()

5       On Error GoTo AfficherErreur

10      mskPMHeureFin.mask = vbNullString

15      If mskPMHeureFin.Text = "__:__" Then
20        mskPMHeureFin.Text = vbNullString
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmPunch", "mskPMHeureFin_LostFocus", Err, Erl
End Sub

Private Sub mskNoProjet_Change()

5       On Error GoTo AfficherErreur

10      If fraPunch.Visible = True Then
15        If InStr(1, mskNoProjet.Text, "_") = 0 Then
20          Call AfficherTypePunch

25          Call AfficherClient
30        Else
35          txtClient.Text = vbNullString
40        End If
45      End If

50      Call RemplirComboType

55      Exit Sub

AfficherErreur:

60      woups "frmPunch", "mskNoProjet_Change", Err, Erl
End Sub

Private Sub mskPMNoProjet_Change()

5       On Error GoTo AfficherErreur

10      If fraPunchMultiple.Visible = True Then
15        If InStr(1, mskPMNoProjet.Text, "_") = 0 Then
20          Call AfficherTypePunch

25          Call AfficherClient
30        Else
35          txtPMClient.Text = vbNullString
40        End If
45      End If

50      Call RemplirComboType

55      Exit Sub

AfficherErreur:

60      woups "frmPunch", "mskPMNoProjet_Change", Err, Erl
End Sub

Private Sub AfficherTypePunch()
  
5       On Error GoTo AfficherErreur

10      Dim sNumero As String
15      Dim sType   As String
20      Dim bPM     As Boolean
  
25      If fraPunchMultiple.Visible = True Then
30        sNumero = mskPMNoProjet.Text
35        bPM = True
40      Else
45        If mskNoProjet.Text <> "_____-__" Then
50          sNumero = mskNoProjet.Text
55        Else
60          sNumero = txtnoprojet.Text
61        End If

62        bPM = False
63      End If
  
64      If Left$(sNumero, 5) = Right$(Year(Date), 1) & "3000" Then
65        Select Case Right$(sNumero, 2)
            Case "60": sType = "Petits outils && fourniture"
70          Case "70": sType = "Administration de la shop"
75          Case "71": sType = "Identification de fils, lamicoïdes, etc."
80          Case "72": sType = "Réception de marchandise"
85          Case "73": sType = "Support technique informatique et téléphone"
90          Case "74": sType = "Commissions"
95          Case "75": sType = "Site web && publications"
100         Case "76": sType = "Entretien && réparation de la bâtisse"
105         Case "77": sType = "Ménage général"
110         Case "80": sType = "Réparation des outils GRB"
115         Case "81": sType = "Lavage des véhicules"
120         Case "82": sType = "Entretien && réparation véhicules"
125         Case "83": sType = "Formation du personnel"
130         Case "85": sType = "Logiciel interne"
135         Case "95": sType = "Bâtiment"
140         Case "97": sType = "Équipement bureau && informatique"
145         Case "99": sType = "Équipements && outillage"
150         Case Else: sType = vbNullString
155       End Select

160       If bPM = True Then
165         lblPMTypePunch.Caption = sType
170       Else
175         lblTypePunch.Caption = sType
180       End If
185     Else
190       If bPM = True Then
195         lblPMTypePunch.Caption = vbNullString
200       Else
205         lblTypePunch.Caption = vbNullString
210       End If
215     End If

220     Exit Sub

AfficherErreur:

225     woups "frmPunch", "AfficherTypePunch", Err, Erl
End Sub

Private Sub AfficherClient()

5       On Error GoTo AfficherErreur

10      Dim rstProjSoum As ADODB.Recordset
15      Dim rstClient   As ADODB.Recordset
20      Dim iCompteur   As Integer
25      Dim sPrefixe    As String

30      If fraPunchMultiple.Visible = True Then
35        If optPMTypePunch(I_OPT_ELECTRIQUE).Value = True Then
40          sPrefixe = "E"
45        Else
50          sPrefixe = "M"
55        End If
60      Else
65        If optTypePunch(I_OPT_ELECTRIQUE).Value = True Then
70          sPrefixe = "E"
75        Else
80          sPrefixe = "M"
85        End If
90      End If
    
95      Set rstProjSoum = New ADODB.Recordset
100     Set rstClient = New ADODB.Recordset
    
105     If fraPunchMultiple.Visible = True Then
110       Call rstProjSoum.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & sPrefixe & mskPMNoProjet.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
115     Else
120       Call rstProjSoum.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & sPrefixe & mskNoProjet.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
125     End If
  
130     If Not rstProjSoum.EOF Then
135       Call rstClient.Open("SELECT NomClient FROM GRB_Client WHERE IDClient = " & rstProjSoum.Fields("NoClient"), g_connData, adOpenDynamic, adLockOptimistic)
    
140       If fraPunchMultiple.Visible = True Then
145         txtPMClient.Text = rstClient.Fields("NomClient")
150         txtPMClient.Tag = rstProjSoum.Fields("NoClient")
155       Else
160         txtClient.Text = rstClient.Fields("NomClient")
165         txtClient.Tag = rstProjSoum.Fields("NoClient")
170       End If
    
175       Call rstClient.Close
180       Set rstClient = Nothing
    
185       If rstProjSoum.Fields("Ouvert") = False Then
190         If rstProjSoum.Fields("Type") = "P" Then
195           Call MsgBox("Ce projet est fermé!", vbOKOnly, "Erreur")
200         Else
205           Call MsgBox("Cette soumission est fermée!", vbOKOnly, "Erreur")
210         End If
215       End If
220     Else
225       Call MsgBox("Numéro inexistant!", vbOKOnly, "Erreur")

230       txtClient.Text = ""
235       txtClient.Tag = ""
240     End If
  
245     Call rstProjSoum.Close
250     Set rstProjSoum = Nothing

255     Exit Sub

AfficherErreur:

260     woups "frmPunch", "AfficherClient", Err, Erl
End Sub

Private Sub mvwSelection_GotFocus()

'Cette procédure sert à éliminer un bug du controle Active X MonthView
'C'est un bug connu pas Microsoft et la solution suivante est proposée
'Il faut avoir une variable boolean mise à true si le MonthView prend le focus
'et ensuite, en cliquant sur un bouton, si le MonthView a le focus, on force le clique

5       On Error GoTo AfficherErreur

10      m_bMonthViewHasFocus = True

15      Exit Sub

AfficherErreur:

20      woups "frmPunch", "mvwSelection_GotFocus", Err, Erl
End Sub

Private Sub optHeure_Click(Index As Integer)

5       On Error GoTo AfficherErreur

10      If Index = I_OPT_SYSTEME Then
15        mskHeure.Enabled = False
20      Else
25        mskHeure.Enabled = True

30        If m_bModifPunch = False Then
35          Call mskHeure.SetFocus
40        End If
45      End If

50      Exit Sub

AfficherErreur:

55      woups "frmPunch", "optHeure_Click", Err, Erl
End Sub

Private Sub lvwJourSemaine_Click(Index As Integer)

5       On Error GoTo AfficherErreur

10      Dim iCompteur  As Integer
15      Dim sDate      As String
20      Dim iNbreJour  As Integer
    
        'Initialise la couleur en blanc
25      For iCompteur = 1 To 7
30        lvwJourSemaine(iCompteur).BackColor = &HFFFFFF
35      Next
  
        'Sélectionne jour de semaine
40      lvwJourSemaine(Index).BackColor = &HE0E0E0

45      sDate = lvwJourSemaine(Index).Tag

50      Select Case Mid$(sDate, 6, 2)
          Case "01": iNbreJour = 31
55        Case "02":
60          If CInt(Left$(sDate, 4)) Mod 4 = 0 Then
65            iNbreJour = 29
70          Else
75            iNbreJour = 28
80          End If

85        Case "03": iNbreJour = 31
90        Case "04": iNbreJour = 30
95        Case "05": iNbreJour = 31
100       Case "06": iNbreJour = 30
105       Case "07": iNbreJour = 31
110       Case "08": iNbreJour = 31
115       Case "09": iNbreJour = 30
120       Case "10": iNbreJour = 31
125       Case "11": iNbreJour = 30
130       Case "12": iNbreJour = 31
135     End Select

140     Do While mvwSelection.Day >= iNbreJour
145       mvwSelection.Day = mvwSelection.Day - 1
150     Loop

        'Sélectionne dans calendrier
155     mvwSelection.Year = Left$(sDate, 4)
160     mvwSelection.Month = Mid$(sDate, 6, 2)
165     mvwSelection.Day = Right$(sDate, 2)

        'Date choisie
170     m_datDateChoisie = DateSerial(mvwSelection.Year, mvwSelection.Month, mvwSelection.Day)

        'Affiche horaire jour
175     Call RemplirListViewJour
180     Call RemplirListViewJourAutorisation

185     frajour.Visible = True
190     fraPunch.Visible = False

195     Call lvwJour.SetFocus

200     Exit Sub

AfficherErreur:

205     woups "frmPunch", "lvwJourSemaine_Click", Err, Erl, "Date cliquée : " & sDate)
End Sub

Private Sub mskNoProjet_KeyPress(KeyAscii As Integer)

5       On Error GoTo AfficherErreur

        'Pour changer un "m" en un "M"
10      If KeyAscii = 109 Then '109 = m
15        KeyAscii = vbKeyM 'M
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmPunch", "mskNoProjet_KeyPress", Err, Erl
End Sub

Private Sub mskPMNoProjet_KeyPress(KeyAscii As Integer)

5       On Error GoTo AfficherErreur

        'Pour changer un "m" en un "M"
10      If KeyAscii = 109 Then '109 = m
15        KeyAscii = vbKeyM 'M
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmPunch", "mskPMNoProjet_KeyPress", Err, Erl
End Sub

Private Sub cmdOK_Click()

5       On Error GoTo AfficherErreur

        'Enregistrement du punch in ou du punch out
10      Dim rstPunch      As ADODB.Recordset
15      Dim rstProjSoum   As ADODB.Recordset
20      Dim sHeure        As String
25      Dim bModif        As Boolean
30      Dim iCompteur     As Integer
35      Dim sPrefixe      As String
40      Dim sType         As String
45      Dim sNoProjet     As String
50      Dim bInstallation As Boolean
55      Dim sHeureFin     As String
        Dim sNumero As String
        
        
If m_ePunch = I_MODIF_PUNCH_IN Or m_ePunch = I_PUNCH_IN Then
   sNumero = mskNoProjet.Text
    Else

         sNumero = txtnoprojet.Text
     End If
60      m_bMonthViewHasFocus = False

65      If optHeure(I_OPT_SYSTEME).Value = True Then
70        sHeure = GetHeure(Time)
75        bModif = False
80      Else
85        If mskHeure.Text <> vbNullString Then
90          If InStr(1, mskHeure.Text, "_") = 0 Then
95            sHeure = GetHeure(mskHeure.Text)
100           bModif = True
105         Else
110           Call MsgBox("Heure invalide!", vbOKOnly, "Erreur")
      
115           Exit Sub
120         End If
125       Else
130         Call MsgBox("Heure invalide!", vbOKOnly, "Erreur")

135         Exit Sub
140       End If
145     End If
  
150     If sHeure <> "" Then
155       If m_ePunch = I_MODIF_PUNCH_IN Or m_ePunch = I_PUNCH_IN Then
160         If cmbemployé.ListIndex = -1 Or InStr(1, mskNoProjet.Text, "_") > 0 Then
165           Call MsgBox("Le nom de l'employé et le numéro de projet sont des champs obligatoires!", vbOKOnly, "Erreur")
     
170           Exit Sub
175         End If
180       End If

185       If cmbType.ListIndex = -1 And cmbType.Visible = True Then
190         Call MsgBox("Le type est obligatoire!", vbOKOnly, "Erreur")

195         Exit Sub
200       End If
  
          'Si c'est une modification de punch in, il faut vérifier l'heure
          'pour être sur qu'elle sont correctes chronologiquement
205       If m_ePunch = I_MODIF_PUNCH_IN Then
210         If lvwJour.SelectedItem.SubItems(I_LVW_FIN) <> vbNullString Then
215           If sHeure > lvwJour.SelectedItem.SubItems(I_LVW_FIN) Then
220             Call MsgBox("L'heure de début doit être plus petite que l'heure de fin!", vbOKOnly, "Erreur")
         
225             Exit Sub
230           End If
235         End If
240       End If

          'Si c'est une modification de punch in, il faut vérifier l'heure
          'pour être sur qu'elle sont correctes chronologiquement
245       If m_ePunch = I_MODIF_PUNCH_OUT Or m_ePunch = I_PUNCH_OUT Then
250         If sHeure < lvwJour.SelectedItem.SubItems(I_LVW_DEBUT) Then
255           Call MsgBox("L'heure de fin doit être plus grande que l'heure de début!", vbOKOnly, "Erreur")
        
260           Exit Sub
265         End If
270       End If

          'Si c'est un punch out avec l'heure de diner et que c'est avant l'heure du diner
275       If m_ePunch = I_PUNCH_OUT Or m_ePunch = I_MODIF_PUNCH_OUT Then
280         If chkDiner.Value = vbChecked Then
285           If optHeureDiner(I_OPT_30_MINUTES).Value = True Then
290             If sHeure < "12:30" Then
295               Call MsgBox("La case à cocher 'Heure du dîner' ne peut être cochée seulement si l'heure de fin est plus grande que 12:30!", vbOKOnly, "Erreur")

300               Exit Sub
305             Else
310               If lvwJour.SelectedItem.SubItems(I_LVW_DEBUT) > "12:00" Then
315                 Call MsgBox("La case à cocher 'Heure du dîner' ne peut être cochée seulement si l'heure de début est plus petite que 12:00!", vbOKOnly, "Erreur")

320                 Exit Sub
325               End If
330             End If
335           Else
340             If sHeure < "13:00" Then
345               Call MsgBox("La case à cocher 'Heure du dîner' ne peut être cochée seulement si l'heure de fin est plus grande que 13:00!", vbOKOnly, "Erreur")

350               Exit Sub
355             Else
360               If lvwJour.SelectedItem.SubItems(I_LVW_DEBUT) > "12:00" Then
365                 Call MsgBox("La case à cocher 'Heure du dîner' ne peut être cochée seulement si l'heure de début est plus petite que 12:00!", vbOKOnly, "Erreur")

370                 Exit Sub
375               End If
380             End If
385           End If
390         End If
395       End If
  
400       If optTypePunch(I_OPT_ELECTRIQUE).Value = True Then
405         sPrefixe = "E"
410       Else
415         sPrefixe = "M"
420       End If
  
425       If m_ePunch = I_MODIF_PUNCH_IN Or m_ePunch = I_PUNCH_IN Then
430         sNoProjet = mskNoProjet.Text
435       Else
440         sNoProjet = txtnoprojet.Text
445       End If

450       If cmbType.Visible = True Then
455         If IsNumeric(Right$(sNoProjet, 2)) Then
460           If CInt(Right$(sNoProjet, 2)) >= 51 And CInt(Right$(sNoProjet, 2)) <= 59 Then
465             bInstallation = True
470           Else
475             bInstallation = False
480           End If
485         Else
490           bInstallation = False
495         End If
  
500         If bInstallation = True Then
505           If sPrefixe = "E" Then
510             Select Case cmbType.ListIndex
                  Case I_TYPE_ELEC_INSTALLATION: sType = "Installation"
515               Case I_TYPE_ELEC_MISE_SERVICE: sType = "MiseService"
520             End Select
525           Else
530             Select Case cmbType.ListIndex
                  Case I_TYPE_MEC_INSTALLATION: sType = "Installation"
535             End Select
540           End If
545         Else
550           If sPrefixe = "E" Then
555             sType = cmbType.Text
610
615           Else
620
                sType = cmbType.Text
670
675           End If
680         End If
685       End If

690       If m_ePunch = I_MODIF_PUNCH_IN Or m_ePunch = I_PUNCH_IN Then
695         Set rstProjSoum = New ADODB.Recordset

700         Call rstProjSoum.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & sPrefixe & mskNoProjet.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
    
705         If Not rstProjSoum.EOF Then
710           If txtClient.Text <> "" Then
715             If rstProjSoum.Fields("Ouvert") = False Then
720               If rstProjSoum.Fields("Type") = "P" Then
725                 Call MsgBox("Ce projet est fermé!", vbOKOnly, "Erreur")
730               Else
735                 Call MsgBox("Cette soumission est fermée!", vbOKOnly, "Erreur")
740               End If
    
745               Call rstProjSoum.Close
750               Set rstProjSoum = Nothing
       
755               Exit Sub
760             End If
765           Else
770             Call MsgBox("Le client ne doit pas être vide!", vbOKOnly, "Erreur")

775             Call rstProjSoum.Close
780             Set rstProjSoum = Nothing
       
785             Exit Sub
790           End If
795         Else
800           Call MsgBox("Numéro inexistant!", vbOKOnly, "Erreur")
      
805           Call rstProjSoum.Close
810           Set rstProjSoum = Nothing
      
815           Exit Sub
820         End If
            
825         Call rstProjSoum.Close
830         Set rstProjSoum = Nothing
835       End If

840       If Trim$(txtCommentaires.Text) = "" Then
845         Call MsgBox("Le commentaire est obligatoire!", vbOKOnly, "Erreur")

850         Exit Sub
855       End If
                  
860       Set rstPunch = New ADODB.Recordset
                  
          'Selon le mode
865       Select Case m_ePunch
            'Si c'est un punch in
            Case I_PUNCH_IN:
              'On ouvre le recordset avec la date et le no d'employé
870           Call rstPunch.Open("SELECT * FROM GRB_Punch WHERE NoEmploye = " & cmbemployé.ItemData(cmbemployé.ListIndex) & " AND Date = '" & ConvertDate(m_datDateChoisie) & "' ORDER BY Date,HeureDébut", g_connData, adOpenDynamic, adLockOptimistic)
        
              'Si il y a des enregistrements
875           If Not rstPunch.EOF Then
                'On va au dernier
880             Call rstPunch.MoveLast
      
                'On vérifie si le dernier punch out n'a pas été fait
885             If IsNull(rstPunch.Fields("HeureFin")) Or rstPunch.Fields("HeureFin") = vbNullString Then
                  'On fait le punch out
890               rstPunch.Fields("ModifFin") = bModif
895               rstPunch.Fields("HeureFin") = sHeure

900               Call rstPunch.Update
905             End If
910           End If
      
915           Call rstPunch.AddNew
      
920           rstPunch.Fields("NoEmploye") = cmbemployé.ItemData(cmbemployé.ListIndex)
925           rstPunch.Fields("NoProjet") = sPrefixe & mskNoProjet.Text
930           rstPunch.Fields("Date") = ConvertDate(DateSerial(mvwSelection.Year, mvwSelection.Month, mvwSelection.Day))
935           rstPunch.Fields("ModifDébut") = bModif
940           rstPunch.Fields("HeureDébut") = sHeure
945           rstPunch.Fields("Commentaire") = txtCommentaires.Text
950           rstPunch.Fields("NoClient") = txtClient.Tag

955           If chkKM.Value = vbChecked Then
960             rstPunch.Fields("KM") = True

965             If txtKM.Text <> "" Then
970               txtKM.Text = Replace(txtKM.Text, ".", ",")

975               If IsNumeric(txtKM.Text) Then
980                 rstPunch.Fields("NbreKM") = txtKM.Text
985               Else
990                 rstPunch.Fields("NbreKM") = 0
995               End If
1000            Else
1005              rstPunch.Fields("KM") = False
1010              rstPunch.Fields("NbreKM") = ""
1015            End If
1020          Else
1025            rstPunch.Fields("KM") = False
1030            rstPunch.Fields("NbreKM") = ""
1035          End If
            If Mid$(sNumero, 2, 1) = "1" Then
                rstPunch.Fields("Type") = "Soumission"
             Else
1040          rstPunch.Fields("Type") = sType
             End If
1045          Call rstPunch.Update
     
1050          Call rstPunch.Close
          
            'Si c'est un punch out
1055        Case I_PUNCH_OUT:
              'Si l'élément choisi est en noir, le punch out a déjà été fait
1060          If lvwJour.SelectedItem.ForeColor = COLOR_ROUGE Then
                'On ouvre le recordset avec le numéro de punch
1065            Call rstPunch.Open("SELECT * FROM GRB_Punch WHERE IDPunch = " & lvwJour.SelectedItem.Tag, g_connData, adOpenDynamic, adLockOptimistic)
      
1070            If chkDiner.Value = vbChecked Then
1075              rstPunch.Fields("ModifFin") = False
1080              rstPunch.Fields("HeureFin") = "12:00"
1085            Else
1090              rstPunch.Fields("ModifFin") = bModif
1095              rstPunch.Fields("HeureFin") = sHeure
1100            End If

1105            rstPunch.Fields("Commentaire") = txtCommentaires.Text
      
1110            If chkKM.Value = vbChecked Then
1115              rstPunch.Fields("KM") = True

1120              If txtKM.Text <> "" Then
1125                txtKM.Text = Replace(txtKM.Text, ".", ",")

1130                If IsNumeric(txtKM.Text) Then
1135                  rstPunch.Fields("NbreKM") = txtKM.Text
1140                Else
1145                  rstPunch.Fields("NbreKM") = 0
1150                End If
1155              Else
1160                rstPunch.Fields("KM") = False
1165                rstPunch.Fields("NbreKM") = ""
1170              End If
1175            Else
1180              rstPunch.Fields("KM") = False
1185              rstPunch.Fields("NbreKM") = ""
1190            End If
                If Mid$(sNumero, 2, 1) = "1" Then
                    rstPunch.Fields("Type") = "Soumission"
                Else
1195                rstPunch.Fields("Type") = sType
                End If
1200            Call rstPunch.Update

1205            If chkDiner.Value = vbChecked Then
1210              Call rstPunch.AddNew
     
1215              rstPunch.Fields("NoEmploye") = cmbemployé.ItemData(cmbemployé.ListIndex)

1220              If mskNoProjet.Text = "_____-__" Then
1225                rstPunch.Fields("NoProjet") = sPrefixe & txtnoprojet.Text
1230              Else
1235                rstPunch.Fields("NoProjet") = sPrefixe & mskNoProjet.Text
1240              End If
               
1245              rstPunch.Fields("Date") = ConvertDate(DateSerial(mvwSelection.Year, mvwSelection.Month, mvwSelection.Day))
1250              rstPunch.Fields("ModifDébut") = False

1255              If optHeureDiner(I_OPT_30_MINUTES).Value = True Then
1260                rstPunch.Fields("HeureDébut") = "12:30"
1265              Else
1270                rstPunch.Fields("HeureDébut") = "13:00"
1275              End If

1280              rstPunch.Fields("Commentaire") = txtCommentaires.Text
1285              rstPunch.Fields("NoClient") = txtClient.Tag
1290              rstPunch.Fields("ModifFin") = bModif
1295              rstPunch.Fields("HeureFin") = sHeure
                If Mid$(sNumero, 2, 1) = "1" Then
                    rstPunch.Fields("Type") = "Soumission"
                Else
1300              rstPunch.Fields("Type") = sType
                End If
1305              Call rstPunch.Update
1310            End If
      
1315            Call rstPunch.Close
1320          Else
1325            Call MsgBox("Le punch out a déjà été fait!", vbOKOnly, "Erreur")
1330          End If
      
            'Si c'est une modification de punch in
1335        Case I_MODIF_PUNCH_IN:
              'On ouvre le recordset avec le numéro de punch
1340          Call rstPunch.Open("SELECT * FROM GRB_Punch WHERE IDPunch = " & lvwJour.SelectedItem.Tag, g_connData, adOpenDynamic, adLockOptimistic)
        
1345          rstPunch.Fields("NoEmploye") = cmbemployé.ItemData(cmbemployé.ListIndex)
1350          rstPunch.Fields("NoProjet") = sPrefixe & mskNoProjet.Text
1355          rstPunch.Fields("NoClient") = txtClient.Tag
1360          rstPunch.Fields("Date") = ConvertDate(DateSerial(mvwSelection.Year, mvwSelection.Month, mvwSelection.Day))

1365          If bModif = True Then
1370            If rstPunch.Fields("HeureDébut") <> sHeure Then
1375              rstPunch.Fields("ModifDébut") = True
1380            Else
1385              rstPunch.Fields("ModifDébut") = False
1390            End If
1395          Else
1400            rstPunch.Fields("ModifDébut") = False
1405          End If
                
1410          rstPunch.Fields("HeureDébut") = sHeure

1415          rstPunch.Fields("Commentaire") = txtCommentaires.Text
         
1420          If chkKM.Value = vbChecked Then
1425            rstPunch.Fields("KM") = True

1430            If txtKM.Text <> "" Then
1435              txtKM.Text = Replace(txtKM.Text, ".", ",")

1440              If IsNumeric(txtKM.Text) Then
1445                rstPunch.Fields("NbreKM") = txtKM.Text
1450              Else
1455                rstPunch.Fields("NbreKM") = 0
1460              End If
1465            Else
1470              rstPunch.Fields("KM") = False
1475              rstPunch.Fields("NbreKM") = 0
1480            End If
1485          Else
1490            rstPunch.Fields("KM") = False
1495            rstPunch.Fields("NbreKM") = ""
1500          End If
            If Mid$(sNumero, 2, 1) = "1" Then
                rstPunch.Fields("Type") = "Soumission"
             Else
1505          rstPunch.Fields("Type") = sType
             End If
1510          Call rstPunch.Update
  
1515          Call rstPunch.Close
      
            'Si c'est une modification de punch out
1520        Case I_MODIF_PUNCH_OUT:
              'On ouvre le recordset avec le numéro de punch
1525          Call rstPunch.Open("SELECT * FROM GRB_Punch WHERE IDPunch = " & lvwJour.SelectedItem.Tag, g_connData, adOpenDynamic, adLockOptimistic)
    
1530          If chkDiner.Value = vbChecked Then
1535            sHeureFin = rstPunch.Fields("HeureFin")

1540            rstPunch.Fields("ModifFin") = False
1545            rstPunch.Fields("HeureFin") = "12:00"
1550          Else
1555            If bModif = True Then
1560              If rstPunch.Fields("HeureFin") <> sHeure Then
1565                rstPunch.Fields("ModifFin") = True
1570              Else
1575                rstPunch.Fields("ModifFin") = False
1580              End If
1585            Else
1590              rstPunch.Fields("ModifFin") = False
1595            End If

1600            rstPunch.Fields("HeureFin") = sHeure
1605          End If

1610          rstPunch.Fields("Commentaire") = txtCommentaires.Text
    
1615          If chkKM.Value = vbChecked Then
1620            rstPunch.Fields("KM") = True

1625            If txtKM.Text <> "" Then
1630              txtKM.Text = Replace(txtKM.Text, ".", ",")

1635              If IsNumeric(txtKM.Text) Then
1640                rstPunch.Fields("NbreKM") = txtKM.Text
1645              Else
1650                rstPunch.Fields("NbreKM") = 0
1655              End If
1660            Else
1665              rstPunch.Fields("KM") = False
1670              rstPunch.Fields("NbreKM") = ""
1675            End If
1680          Else
1685            rstPunch.Fields("KM") = False
1690            rstPunch.Fields("NbreKM") = ""
1695          End If

            If Mid$(sNumero, 2, 1) = "1" Then
                rstPunch.Fields("Type") = "Soumission"
             Else
1700          rstPunch.Fields("Type") = sType
            End If
1705          Call rstPunch.Update

1710          If chkDiner.Value = vbChecked Then
1715            Call rstPunch.AddNew
    
1720            rstPunch.Fields("NoEmploye") = cmbemployé.ItemData(cmbemployé.ListIndex)

1725            rstPunch.Fields("NoProjet") = sPrefixe & txtnoprojet.Text

1730            rstPunch.Fields("Date") = ConvertDate(DateSerial(mvwSelection.Year, mvwSelection.Month, mvwSelection.Day))
1735            rstPunch.Fields("ModifDébut") = False

1740            If optHeureDiner(I_OPT_30_MINUTES).Value = True Then
1745              rstPunch.Fields("HeureDébut") = "12:30"
1750            Else
1755              rstPunch.Fields("HeureDébut") = "13:00"
1760            End If

1765            rstPunch.Fields("Commentaire") = txtCommentaires.Text
1770            rstPunch.Fields("NoClient") = txtClient.Tag

1775            If bModif = True Then
1780              If rstPunch.Fields("HeureFin") <> sHeureFin Then
1785                rstPunch.Fields("ModifFin") = True
1790              Else
1795                rstPunch.Fields("ModifFin") = False
1800              End If
1805            Else
1810              rstPunch.Fields("ModifFin") = False
1815            End If

1820            rstPunch.Fields("HeureFin") = sHeure

1825            rstPunch.Fields("Type") = sType

1830            Call rstPunch.Update
1835          End If
             
1840          Call rstPunch.Close
1845      End Select
  
1850      Set rstPunch = Nothing
  
1855      Call RemplirListViewSemaine(True)
1860      Call RemplirListViewSemaineAutorisation(True)
1865      Call RemplirListViewJour
1870      Call RemplirListViewJourAutorisation

1875      Call CalculerHeureSemaine

1880      frajour.Visible = True
1885      fraPunch.Visible = False
1890      fraPunchMultiple.Visible = False
1895    End If

1900    Exit Sub

AfficherErreur:

1905    woups "frmPunch", "cmdOK_Click", Err, Erl
End Sub

Private Sub cmdPMOK_Click()

5       On Error GoTo AfficherErreur

        'Enregistrement du punch in ou du punch out
10      Dim rstPunch      As ADODB.Recordset
15      Dim rstProjSoum   As ADODB.Recordset
20      Dim sHeureDebut   As String
25      Dim sHeureFin     As String
30      Dim iCompteur     As Integer
35      Dim bSelect       As Boolean
40      Dim sPrefixe      As String
45      Dim sType         As String
50      Dim bInstallation As Boolean

55      m_bMonthViewHasFocus = False

60      If mskPMHeureDebut.Text <> vbNullString Then
65        If InStr(1, mskPMHeureDebut.Text, "_") = 0 Then
70          sHeureDebut = GetHeure(mskPMHeureDebut.Text)
75        Else
80          Call MsgBox("Heure de début invalide!", vbOKOnly, "Erreur")
      
85          Exit Sub
90        End If
95      Else
100       Call MsgBox("Heure de début invalide!", vbOKOnly, "Erreur")

105       Exit Sub
110     End If
  
115     If mskPMHeureFin.Text <> vbNullString Then
120       If InStr(1, mskPMHeureFin.Text, "_") = 0 Then
125         sHeureFin = GetHeure(mskPMHeureFin.Text)
130       Else
135         Call MsgBox("Heure de fin invalide!", vbOKOnly, "Erreur")
      
140         Exit Sub
145       End If
150     Else
155       Call MsgBox("Heure de fin invalide!", vbOKOnly, "Erreur")

160       Exit Sub
165     End If

170     If cmbPMType.ListIndex = -1 And cmbPMType.Visible = True Then
175       Call MsgBox("Le type est obligatoire!", vbOKOnly, "Erreur")

180       Exit Sub
185     End If

190     For iCompteur = 1 To lvwEmployes.ListItems.count
195       If lvwEmployes.ListItems(iCompteur).Checked = True Then
200         bSelect = True

205         Exit For
210       End If
215     Next

220     If bSelect = False Then
225       Call MsgBox("Vous devez choisir au moins 1 employé!", vbOKOnly, "Erreur")

230       Exit Sub
235     End If

240     If InStr(1, mskPMNoProjet.Text, "_") > 0 Then
245       Call MsgBox("Le numéro de projet est obligatoire!", vbOKOnly, "Erreur")
   
250       Exit Sub
255     End If

260     If sHeureDebut > sHeureFin Then
265       Call MsgBox("L'heure de début doit être plus petite que l'heure de fin!", vbOKOnly, "Erreur")
       
270       Exit Sub
275     End If

        'Si c'est un punch out avec l'heure de diner et que c'est avant l'heure du diner
280     If chkPMDiner.Value = vbChecked Then
285       If optPMHeureDiner(I_OPT_30_MINUTES).Value = True Then
290         If sHeureDebut > "12:00" Or sHeureFin < "12:30" Then
295           Call MsgBox("La case à cocher 'Heure du dîner' ne peut être cochée que si l'heure de début est plus petite que 12:00" & vbNewLine & _
                          " ou que l'heure de fin est plus grande que 12:30!", vbOKOnly, "Erreur")

300           Exit Sub
305         End If
310       Else
315         If sHeureDebut > "12:00" Or sHeureFin < "13:00" Then
320           Call MsgBox("La case à cocher 'Heure du dîner' ne peut être cochée que si l'heure de début est plus petite que 12:00" & vbNewLine & _
                          " ou que l'heure de fin est plus grande que 13:00!", vbOKOnly, "Erreur")

325           Exit Sub
330         End If
335       End If
340     End If

345     If optPMTypePunch(I_OPT_ELECTRIQUE).Value = True Then
350       sPrefixe = "E"
355     Else
360       sPrefixe = "M"
365     End If
        
370     If cmbPMType.Visible = True Then
375       If IsNumeric(Right$(mskPMNoProjet.Text, 2)) Then
380         If CInt(Right$(mskPMNoProjet.Text, 2)) >= 51 And CInt(Right$(mskPMNoProjet.Text, 2)) <= 59 Then
385           bInstallation = True
390         Else
395           bInstallation = False
400         End If
405       Else
410         bInstallation = False
415       End If
        
420       If bInstallation = True Then
425         If sPrefixe = "E" Then
430           Select Case cmbPMType.ListIndex
                Case I_TYPE_ELEC_INSTALLATION: sType = "Installation"
435             Case I_TYPE_ELEC_MISE_SERVICE: sType = "Formation"
440           End Select
445         Else
450           Select Case cmbPMType.ListIndex
                Case I_TYPE_MEC_INSTALLATION: sType = "Installation"
455           End Select
460         End If
465       Else
470         If sPrefixe = "E" Then
                sType = cmbPMType.Text
480
530
535         Else
540           sType = cmbPMType.Text

590
595         End If
600       End If
605     End If

610     Set rstProjSoum = New ADODB.Recordset

615     Call rstProjSoum.Open("SELECT * FROM GRB_ProjSoum WHERE IDProjSoum = '" & sPrefixe & mskPMNoProjet.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
  
620     If Not rstProjSoum.EOF Then
625       If txtPMClient.Text <> "" Then
630         If rstProjSoum.Fields("Ouvert") = False Then
635           If rstProjSoum.Fields("Type") = "P" Then
640             Call MsgBox("Ce projet est fermé!", vbOKOnly, "Erreur")
645           Else
650             Call MsgBox("Cette soumission est fermée!", vbOKOnly, "Erreur")
655           End If
  
660           Call rstProjSoum.Close
665           Set rstProjSoum = Nothing
    
670           Exit Sub
675         End If
680       Else
685         Call MsgBox("Le client ne doit pas être vide!", vbOKOnly, "Erreur")

690         Call rstProjSoum.Close
695         Set rstProjSoum = Nothing

700         Exit Sub
705       End If
710     Else
715       Call MsgBox("Numéro inexistant!", vbOKOnly, "Erreur")

720       Call rstProjSoum.Close
725       Set rstProjSoum = Nothing
    
730       Exit Sub
735     End If
       
740     Call rstProjSoum.Close
745     Set rstProjSoum = Nothing

750     If Trim$(txtPMCommentaire.Text) = "" Then
755       Call MsgBox("Le commentaire est obligatoire!", vbOKOnly, "Erreur")

760       Exit Sub
765     End If
                
770     Set rstPunch = New ADODB.Recordset
                 
775     For iCompteur = 1 To lvwEmployes.ListItems.count
780       If lvwEmployes.ListItems(iCompteur).Checked = True Then
785         Call rstPunch.Open("SELECT * FROM GRB_Punch", g_connData, adOpenDynamic, adLockOptimistic)
      
790         Call rstPunch.AddNew
              
795         rstPunch.Fields("NoEmploye") = lvwEmployes.ListItems(iCompteur).Tag
800         rstPunch.Fields("NoProjet") = sPrefixe & mskPMNoProjet.Text
805         rstPunch.Fields("Date") = ConvertDate(DateSerial(mvwSelection.Year, mvwSelection.Month, mvwSelection.Day))
810         rstPunch.Fields("ModifDébut") = False
815         rstPunch.Fields("HeureDébut") = sHeureDebut

820         rstPunch.Fields("ModifFin") = False

825         If chkPMDiner.Value = vbChecked Then
830           rstPunch.Fields("HeureFin") = "12:00"
835         Else
840           rstPunch.Fields("HeureFin") = sHeureFin
845         End If

850         rstPunch.Fields("Commentaire") = txtPMCommentaire.Text
855         rstPunch.Fields("NoClient") = txtPMClient.Tag

860         rstPunch.Fields("KM") = False
865         rstPunch.Fields("NbreKM") = ""

870         rstPunch.Fields("Type") = sType
          
875         Call rstPunch.Update

880         If chkPMDiner.Value = vbChecked Then
885           Call rstPunch.AddNew
    
890           rstPunch.Fields("NoEmploye") = lvwEmployes.ListItems(iCompteur).Tag
895           rstPunch.Fields("NoProjet") = sPrefixe & mskPMNoProjet.Text
900           rstPunch.Fields("Date") = ConvertDate(DateSerial(mvwSelection.Year, mvwSelection.Month, mvwSelection.Day))
905           rstPunch.Fields("ModifDébut") = False

910           If optPMHeureDiner(I_OPT_30_MINUTES).Value = True Then
915             rstPunch.Fields("HeureDébut") = "12:30"
920           Else
925             rstPunch.Fields("HeureDébut") = "13:00"
930           End If

935           rstPunch.Fields("Commentaire") = txtPMCommentaire.Text
940           rstPunch.Fields("NoClient") = txtPMClient.Tag
945           rstPunch.Fields("ModifFin") = False
950           rstPunch.Fields("HeureFin") = sHeureFin

955           rstPunch.Fields("Type") = sType

960           Call rstPunch.Update
965         End If
      
970         Call rstPunch.Update
    
975         Call rstPunch.Close
980       End If
985     Next

990     Set rstPunch = Nothing

995     Call RemplirListViewSemaine(True)
1000    Call RemplirListViewSemaineAutorisation(True)
1005    Call RemplirListViewJour
1010    Call RemplirListViewJourAutorisation

1015    Call CalculerHeureSemaine

1020    frajour.Visible = True
1025    fraPunch.Visible = False
1030    fraPunchMultiple.Visible = False

1035    Exit Sub

AfficherErreur:

1040    woups "frmPunch", "cmdPMOK_Click", Err, Erl
End Sub

Private Function GetHeure(ByVal sHeure As String) As String

5       On Error GoTo AfficherErreur

10      Dim datHeure As Date
15      Dim b24Heure As Boolean

20      If IsNumeric(Left$(sHeure, 2)) And IsNumeric(Mid$(sHeure, 4, 2)) Then
25        If (CInt(Left$(sHeure, 2)) < 0 Or CInt(Left$(sHeure, 2)) > 24) Or (CInt(Mid$(sHeure, 4, 2)) < 0 Or CInt(Mid$(sHeure, 4, 2)) > 59) Then
30          Call MsgBox("Heure incorrecte!", vbOKOnly, "Erreur")

35          Exit Function
40        End If
45      Else
50        Call MsgBox("Heure incorrecte!", vbOKOnly, "Erreur")

55        Exit Function
60      End If

65      sHeure = Replace(sHeure, "AM", "")

70      If InStr(1, sHeure, "PM") > 0 Then
75        sHeure = Trim$(Replace(sHeure, "PM", ""))

80        sHeure = CInt(Left$(sHeure, 2)) + 12 & Right$(sHeure, Len(sHeure) - 2)

85        b24Heure = True
90      End If

95      sHeure = Left$(sHeure, 5)

100     If sHeure <> "24:00" Then
105       datHeure = CDate(sHeure)

110       If Minute(datHeure) <= 5 Then
115         datHeure = TimeSerial(Hour(datHeure), 0, 0)
120       Else
125         If Minute(datHeure) <= 24 Then
130           datHeure = TimeSerial(Hour(datHeure), 15, 0)
135         Else
140           If Minute(datHeure) <= 35 Then
145             datHeure = TimeSerial(Hour(datHeure), 30, 0)
150           Else
155             If Minute(datHeure) <= 54 Then
160               datHeure = TimeSerial(Hour(datHeure), 45, 0)
165             Else
170               datHeure = TimeSerial(Hour(datHeure) + 1, 0, 0)
175             End If
180           End If
185         End If
190       End If

195       GetHeure = Right$("0" & Hour(datHeure), 2) & ":" & Right$("0" & Minute(datHeure), 2)
200     Else
205       GetHeure = sHeure
210     End If

215     Exit Function

AfficherErreur:

220     woups "frmPunch", "GetHeure", Err, Erl
End Function

Private Sub optHeure_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

5       On Error GoTo AfficherErreur

10      If m_bMonthViewHasFocus = True Then
15        If optHeure(Index).Value = False Then
20          optHeure(Index).Value = True
25        End If
30      End If

35      Exit Sub

AfficherErreur:

40      woups "frmPunch", "optHeure_MouseUp", Err, Erl
End Sub

Private Sub optTypePunch_Click(Index As Integer)
        
5       On Error GoTo AfficherErreur

10      If InStr(1, mskNoProjet.Text, "_") = 0 Then
15        Call AfficherTypePunch

20        Call AfficherClient
25      Else
30        If fraPunch.Visible = True Then
35          Call mskNoProjet.SetFocus
40        End If
45      End If

50      Select Case Index
          Case I_OPT_ELECTRIQUE: lblPrefixe.Caption = "E"
55        Case I_OPT_MECANIQUE:  lblPrefixe.Caption = "M"
60      End Select

65      Call RemplirComboType

70      m_bMonthViewHasFocus = False

75      Exit Sub

AfficherErreur:

80      woups "frmPunch", "optTypePunch_Click", Err, Erl
End Sub

Private Sub optTypePunch_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

5       On Error GoTo AfficherErreur

10      If m_bMonthViewHasFocus = True Then
15        Call optTypePunch_Click(Index)
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmPunch", "optTypePunch_MouseUp", Err, Erl
End Sub

Private Sub optPMTypePunch_Click(Index As Integer)
        
5       On Error GoTo AfficherErreur

10      If InStr(1, mskPMNoProjet.Text, "_") = 0 Then
15        Call AfficherTypePunch

20        Call AfficherClient
25      Else
30        If fraPunchMultiple.Visible = True Then
35          Call mskPMNoProjet.SetFocus
40        End If
45      End If

50      Select Case Index
          Case I_OPT_ELECTRIQUE: lblPMPrefixe.Caption = "E"
55        Case I_OPT_MECANIQUE:  lblPMPrefixe.Caption = "M"
60      End Select

65      Call RemplirComboType

70      m_bMonthViewHasFocus = False

75      Exit Sub

AfficherErreur:

80      woups "frmPunch", "optPMTypePunch_Click", Err, Erl
End Sub

Private Sub RemplirComboType()
  
5       On Error GoTo AfficherErreur

10      Dim cmbSource     As ComboBox
15      Dim lblSource     As Label
20      Dim sType         As String
25      Dim sNumero       As String
30      Dim bInstallation As Boolean
        Dim tblremplircombo As ADODB.Recordset
        Set tblremplircombo = New ADODB.Recordset
        
35      If fraPunchMultiple.Visible = True Then
        
40        Set cmbSource = cmbPMType
45        Set lblSource = lblPMType
    
50        sType = lblPMPrefixe.Caption

55        sNumero = mskPMNoProjet.Text
60      Else
        
65        Set cmbSource = cmbType
70        Set lblSource = lblType
    
75        sType = lblPrefixe.Caption

80        If m_ePunch = I_MODIF_PUNCH_IN Or m_ePunch = I_PUNCH_IN Then

85          sNumero = mskNoProjet.Text
90        Else

95          sNumero = txtnoprojet.Text
100       End If
105     End If
  
110     Call cmbSource.Clear

115     If Mid$(sNumero, 2, 1) = "1" Or Mid$(sNumero, 2, 4) = "3000" Then
120       cmbSource.Visible = False
125       lblSource.Visible = False

130       Exit Sub
135     Else

140       cmbSource.Visible = True
145       lblSource.Visible = True
150     End If

155     If IsNumeric(Right$(sNumero, 2)) Then

160       If CInt(Right$(sNumero, 2)) >= 51 And CInt(Right$(sNumero, 2)) <= 59 Then

165         bInstallation = True
170       Else

175         bInstallation = False
180       End If
185     Else

190       bInstallation = False
195     End If
  
200     If bInstallation = True Then

205       If sType = "E" Then

210         Call cmbSource.AddItem("Installation")
215         Call cmbSource.AddItem("Mise en service")
220       Else

225         Call cmbSource.AddItem("Installation")
230       End If
235     Else
240       If sType = "E" Then

245         Call tblremplircombo.Open("Select * from TBL_Punch_Type WHERE MODE = 'E' Order by name ", g_connData, adOpenDynamic, adLockOptimistic)
260         Do While Not tblremplircombo.EOF
265             cmbSource.AddItem (tblremplircombo.Fields("name"))
270         Call tblremplircombo.MoveNext
275
280         Loop
285         Call tblremplircombo.Close
290         Set tblremplircombo = Nothing
295
            
300       Else

305        Call tblremplircombo.Open("Select * from TBL_Punch_Type WHERE MODE = 'M' Order by name ", g_connData, adOpenDynamic, adLockOptimistic)
310        Do While Not tblremplircombo.EOF
315           cmbSource.AddItem (tblremplircombo.Fields("name"))
320           Call tblremplircombo.MoveNext
325        Loop
330        Call tblremplircombo.Close
335        Set tblremplircombo = Nothing
340
345
350
            
355       End If
360     End If
    
365     Set cmbSource = Nothing

370     Exit Sub

AfficherErreur:

375     woups "frmPunch", "RemplirComboType", Err, Erl
End Sub

Private Sub optPMTypePunch_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

5       On Error GoTo AfficherErreur

10      If m_bMonthViewHasFocus = True Then
15        Call optPMTypePunch_Click(Index)
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmPunch", "optPMTypePunch_MouseUp", Err, Erl
End Sub

Private Sub txtKM_LostFocus()
  
5       On Error GoTo AfficherErreur

10      txtKM.Text = Replace(txtKM.Text, ".", ",")

15      Exit Sub
  
AfficherErreur:

20      woups "frmPunch", "txtKM_LostFocus", Err, Erl
End Sub

Private Sub mvwSelection_SelChange(ByVal StartDate As Date, ByVal EndDate As Date, Cancel As Boolean)

5       On Error GoTo AfficherErreur

10      If Month(m_datDateChoisie) <> mvwSelection.Month Or _
          Year(m_datDateChoisie) <> mvwSelection.Year Or _
          Day(m_datDateChoisie) <> mvwSelection.Day Then
15        Call AfficherDate

20        Call CalculerHeureSemaine
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmPunch", "mvwSelection_SelChange", Err, Erl
End Sub

Private Function VerifierModificationDate() As Boolean
        
5       On Error GoTo AfficherErreur

10      Dim bModif              As Boolean
15      Dim datSelected         As Date
20      Dim datToday            As Date
25      Dim datFirstDaySelected As Date
30      Dim datFirstDayToday    As Date

35      datSelected = mvwSelection.Value
40      datToday = Date
45      datFirstDaySelected = GetFirstDay(datSelected)
50      datFirstDayToday = GetFirstDay(datToday)

55      If g_bPunchSemaineAnterieure = False Then
60        If datSelected <> datToday Then
65          If Weekday(datToday, vbSunday) = vbSunday Or _
               Weekday(datToday, vbSunday) = vbMonday Then
70            If (datFirstDaySelected = datFirstDayToday) Or DateDiff("d", datFirstDaySelected, datFirstDayToday) = 7 Then
75              bModif = True
80            Else
85              bModif = False
90            End If
95          Else
100           If datFirstDaySelected = datFirstDayToday Then
105             bModif = True
110           End If
115         End If
120       Else
125         bModif = True
130       End If
135     Else
140       bModif = True
145     End If

150     If bModif = False Then
155       Call MsgBox("Impossible de modifier les punchs de cette journée!", vbOKOnly, "Erreur")
160     End If

165     VerifierModificationDate = bModif

170     Exit Function

AfficherErreur:

175     woups "frmPunch", "VerifierModificationDate", Err, Erl
End Function

Public Sub Table_exist()

Dim adoxconnection As adox.Catalog
Dim i As Integer
Set adoxconnection = New adox.Catalog

adoxconnection.ActiveConnection = g_connData
For i = 0 To adoxconnection.Tables.count - 1
      If LCase(adoxconnection.Tables(i).Name) = LCase("TBL_Punch_Type") Then
        Set adoxconnection = Nothing
        Exit Sub
    End If
Next
Call g_connData.Execute("Create TABLE TBL_Punch_Type (Mode text(1), Name Text (100))")
Call g_connData.Execute("Insert into TBL_Punch_Type (mode, Name) Values ('E','Dessin Électrique');")
Call g_connData.Execute("Insert into TBL_Punch_Type (mode, Name) Values ('E','Fabrication');")
Call g_connData.Execute("Insert into TBL_Punch_Type (mode, Name) Values ('E','Assemblage du Panneau');")
Call g_connData.Execute("Insert into TBL_Punch_Type (mode, Name) Values ('E','Programmation d''interface');")
Call g_connData.Execute("Insert into TBL_Punch_Type (mode, Name) Values ('E','Programmation d''automate');")
Call g_connData.Execute("Insert into TBL_Punch_Type (mode, Name) Values ('E','Programmation de Robot');")
Call g_connData.Execute("Insert into TBL_Punch_Type (mode, Name) Values ('E','Vision Artificielle');")
Call g_connData.Execute("Insert into TBL_Punch_Type (mode, Name) Values ('E','Test Finaux');")
Call g_connData.Execute("Insert into TBL_Punch_Type (mode, Name) Values ('E','Formation du personnel');")
Call g_connData.Execute("Insert into TBL_Punch_Type (mode, Name) Values ('E','Gestion du projet');")
Call g_connData.Execute("Insert into TBL_Punch_Type (mode, Name) Values ('E','Expédition');")
Call g_connData.Execute("Insert into TBL_Punch_Type (mode, Name) Values ('E','Prototypage-Dévelloppement expérimental');")
Call g_connData.Execute("Insert into TBL_Punch_Type (mode, Name) Values ('M','Conception et dessin');")
Call g_connData.Execute("Insert into TBL_Punch_Type (mode, Name) Values ('M','Coupe et préparation(sauf soudage)');")
Call g_connData.Execute("Insert into TBL_Punch_Type (mode, Name) Values ('M','Machinage');")
Call g_connData.Execute("Insert into TBL_Punch_Type (mode, Name) Values ('M','Coupe,Soudure et meulage');")
Call g_connData.Execute("Insert into TBL_Punch_Type (mode, Name) Values ('M','Assemblage des systèmes');")
Call g_connData.Execute("Insert into TBL_Punch_Type (mode, Name) Values ('M','Peinture et Finition');")
Call g_connData.Execute("Insert into TBL_Punch_Type (mode, Name) Values ('M','Test Finaux');")
Call g_connData.Execute("Insert into TBL_Punch_Type (mode, Name) Values ('M','Formation du personnel');")
Call g_connData.Execute("Insert into TBL_Punch_Type (mode, Name) Values ('M','Gestion du projet');")
Call g_connData.Execute("Insert into TBL_Punch_Type (mode, Name) Values ('M','Expédition');")
Call g_connData.Execute("Insert into TBL_Punch_Type (mode, Name) Values ('M','Prototypage-Dévelloppement expérimental');")
Call g_connData.Execute("Insert into TBL_Punch_Type (mode, Name) Values ('S','Soumission');")
Set adoxconnection = Nothing
    

End Sub


