VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPunch 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Punch"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8760
   ScaleWidth      =   13980
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
         ItemData        =   "frmPunch.frx":0000
         Left            =   2760
         List            =   "frmPunch.frx":0002
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
         _ExtentX        =   3254
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
         _ExtentX        =   3254
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
      StartOfWeek     =   152633345
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
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre d'heures dans la semaine pour :"
      ForeColor       =   &H80000008&
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
Private Const I_OPT_SYSTEME As Integer = 0
Private Const I_OPT_MANUELLEMENT As Integer = 1

'Index de optTypePunch et optPMTypePunch
Private Const I_OPT_ELECTRIQUE As Integer = 0
Private Const I_OPT_MECANIQUE As Integer = 1

'Types quand c'est un 51
Private Const I_TYPE_ELEC_INSTALLATION As Integer = 0
Private Const I_TYPE_ELEC_MISE_SERVICE As Integer = 1

'Types quand c'est pas un 51
Private Const I_TYPE_ELEC_DESSIN As Integer = 0
Private Const I_TYPE_ELEC_FABRICATION As Integer = 1
Private Const I_TYPE_ELEC_ASSEMBLAGE As Integer = 2
Private Const I_TYPE_ELEC_PROG_INTERFACE As Integer = 3
Private Const I_TYPE_ELEC_PROG_AUTOMATE As Integer = 4
Private Const I_TYPE_ELEC_PROG_ROBOT As Integer = 5
Private Const I_TYPE_ELEC_VISION As Integer = 6
Private Const I_TYPE_ELEC_TEST As Integer = 7
Private Const I_TYPE_ELEC_FORMATION As Integer = 8
Private Const I_TYPE_ELEC_GESTION As Integer = 9
Private Const I_TYPE_ELEC_SHIPPING As Integer = 10
Private Const I_TYPE_ELEC_prototypage As Integer = 11

'Types quand c'est un 51
Private Const I_TYPE_MEC_INSTALLATION As Integer = 0

'Types quand c'est pas un 51
Private Const I_TYPE_MEC_DESSIN As Integer = 0
Private Const I_TYPE_MEC_COUPE As Integer = 1
Private Const I_TYPE_MEC_MACHINAGE As Integer = 2
Private Const I_TYPE_MEC_SOUDURE As Integer = 3
Private Const I_TYPE_MEC_ASSEMBLAGE As Integer = 4
Private Const I_TYPE_MEC_PEINTURE As Integer = 5
Private Const I_TYPE_MEC_TEST As Integer = 6
Private Const I_TYPE_MEC_FORMATION As Integer = 7
Private Const I_TYPE_MEC_GESTION As Integer = 8
Private Const I_TYPE_MEC_SHIPPING As Integer = 9
Private Const I_TYPE_MEC_prototypage As Integer = 10

'Index de optHeureDiner
Private Const I_OPT_30_MINUTES As Integer = 0
Private Const I_OPT_1_HEURE As Integer = 1

'Index de lvwJour
Private Const I_LVW_NOM As Integer = 0
Private Const I_LVW_PROJET As Integer = 1
Private Const I_LVW_DEBUT As Integer = 2
Private Const I_LVW_FIN As Integer = 3
Private Const I_LVW_CLIENT As Integer = 4
Private Const I_LVW_TYPE As Integer = 5
Private Const I_LVW_COMMENTAIRE As Integer = 6
Private Const I_LVW_KM As Integer = 7

'Index de lvwJourSemaine
Private Const I_LVW_INITIALE As Integer = 0
Private Const I_LVW_TEMPS As Integer = 1

Private Enum enumPunch
 I_PUNCH_IN = 0
 I_PUNCH_OUT = 1
 I_MODIF_PUNCH_IN = 2
 I_MODIF_PUNCH_OUT = 3
End Enum

Private m_ePunch As enumPunch
Private m_iNoEmploye As Integer
Private m_datDateChoisie As Date

Private m_bModifPunch As Boolean
Private m_bMonthViewHasFocus As Boolean

Public sCommentaire As String

Public Sub Afficher(ByVal sUserID As String)

 On Error GoTo Oups

 Dim rstEmploye As ADODB.Recordset
 Dim iCompteur As Integer

 Call Unload(frmChoixPunch)

 Set rstEmploye = New ADODB.Recordset

 Call rstEmploye.Open("SELECT NoEmploye FROM GrbEmployés WHERE loginname = '" & sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)

 m_iNoEmploye = rstEmploye.Fields("NoEmploye")

 Call rstEmploye.Close
 Set rstEmploye = Nothing

 optHeure(I_OPT_SYSTEME).Value = True

 mvwSelection.Year = Year(Date)
  mvwSelection.Month = Month(Date)
  mvwSelection.Day = Day(Date)

  Call AfficherDate

  Call RemplirComboEmploye

  Call cmbHeureSemaine.Clear

  For iCompteur = 0 To cmbemployé.ListCount - 1
  Call cmbHeureSemaine.AddItem(cmbemployé.LIST(iCompteur))

  cmbHeureSemaine.ItemData(cmbHeureSemaine.newIndex) = cmbemployé.ItemData(iCompteur)
10 Next

cmbHeureSemaine.ListIndex = 0

Call Me.Show

Exit Sub

Oups:

wOups "frmPunch", "Afficher", Err, Err.number, Err.Description
End Sub

Private Sub CalculerHeureSemaine()

 On Error GoTo Oups

 Dim rstPunch As ADODB.Recordset
 Dim dblDebut As Double
 Dim dblFin As Double
 Dim dblTotal As Double
 Dim sDate As String
 Dim sDebut As String
 Dim sFin As String

 Set rstPunch = New ADODB.Recordset

 Call rstPunch.Open("SELECT Date, HeureDébut, HeureFin FROM GrbPunch WHERE NoEmploye = " & cmbHeureSemaine.ItemData(cmbHeureSemaine.ListIndex) & " AND Date BETWEEN '" & lvwJourSemaine(1).Tag & "' AND '" & lvwJourSemaine(7).Tag & "'", g_connData, adOpenForwardOnly, adLockReadOnly)

 Do While Not rstPunch.EOF
  sDate = rstPunch.Fields("Date")

  If Not IsNull(rstPunch.Fields("HeureDébut")) Then
  If Trim(rstPunch.Fields("HeureDébut")) <> "" Then
  sDebut = rstPunch.Fields("HeureDébut")
  Else
  sDebut = ""
  End If
Else
sDebut = ""
 End If

 If Not IsNull(rstPunch.Fields("HeureFin")) Then
 If Trim(rstPunch.Fields("HeureFin")) <> "" Then
 sFin = rstPunch.Fields("HeureFin")
 Else
 sFin = ""
 End If
 Else
 sFin = ""
 End If

If sDebut <> "" And sFin <> "" Then
 dblDebut = CDbl(Left$(sDebut, 2)) + CDbl(CDbl(Right$(sDebut, 2)) / 60)
 dblFin = CDbl(Left$(sFin, 2)) + CDbl(CDbl(Right$(sFin, 2)) / 60)

 dblTotal = dblTotal + (dblFin - dblDebut)
 End If

 Call rstPunch.MoveNext
 Loop

1  Call rstPunch.Close
 Set rstPunch = Nothing

 lblNbreHeure.Caption = dblTotal

Exit Sub

Oups:

wOups "frmPunch", "CalculerHeureSemaine", Err, Err.number, Err.Description
End Sub

Private Sub AfficherDate()

 On Error GoTo Oups

 'Affiche punch de la journée et de la semaine
 'dépendant la sélection dans le calendrier
 Dim iCompteur As Integer

 'date choisie
 m_datDateChoisie = DateSerial(mvwSelection.Year, mvwSelection.Month, mvwSelection.Day)

 'affiche punch jour et semaine
 Call RemplirListViewJour
 Call RemplirListViewJourAutorisation
 Call RemplirListViewSemaine(False)
 Call RemplirListViewSemaineAutorisation(False)

 'selectionne jour de la semaine
 For iCompteur = 1 To 7
 If lvwJourSemaine(iCompteur).Tag = m_datDateChoisie Then
 lvwJourSemaine(iCompteur).BackColor = &HE0E0E0
 Else
  lvwJourSemaine(iCompteur).BackColor = &HFFFFFF
  End If
  Next

 'Affiche cedule une journee
  frajour.Visible = True
  fraPunch.Visible = False

  Exit Sub

Oups:

  wOups "frmPunch", "AfficherDate", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListViewJour()

 On Error GoTo Oups

 'remplis ListView une journée
 Dim rstPunch As ADODB.Recordset
 Dim rstEmploye As ADODB.Recordset
 Dim rstClient As ADODB.Recordset
 Dim itmPunch As ListItem
 Dim lForeColor As Long

 'vide le lister
 Call lvwJour.ListItems.Clear

 Set rstPunch = New ADODB.Recordset

 Call rstPunch.Open("SELECT * FROM GrbPunch WHERE Date = '" & ConvertDate(m_datDateChoisie) & "' AND NoEmploye = " & m_iNoEmploye & " ORDER BY HeureDébut", g_connData, adOpenDynamic, adLockOptimistic)

 Set rstEmploye = New ADODB.Recordset
 Set rstClient = New ADODB.Recordset

 'tant il y a de employé cedulé , ajoute dans lister
  Do While Not rstPunch.EOF
  Set itmPunch = lvwJour.ListItems.Add

  Call rstEmploye.Open("SELECT initiale FROM GrbEmployés WHERE NoEmploye = " & m_iNoEmploye, g_connData, adOpenDynamic, adLockOptimistic)

  itmPunch.Text = rstEmploye.Fields("initiale")

  Call rstEmploye.Close

  itmPunch.Tag = rstPunch.Fields("IDPunch")

  If Not IsNull(rstPunch.Fields("NoProjet")) Then
  itmPunch.SubItems(I_LVW_PROJET) = rstPunch.Fields("NoProjet")
Else
itmPunch.SubItems(I_LVW_PROJET) = vbNullString
 End If

 If Not IsNull(rstPunch.Fields("HeureDébut")) Then
 itmPunch.SubItems(I_LVW_DEBUT) = rstPunch.Fields("HeureDébut")
 Else
 itmPunch.SubItems(I_LVW_DEBUT) = vbNullString
 End If

 If Not IsNull(rstPunch.Fields("HeureFin")) And rstPunch.Fields("HeureFin") <> vbNullString Then
 itmPunch.SubItems(I_LVW_FIN) = rstPunch.Fields("HeureFin")
 lForeColor = COLOR_NOIR
 Else
 itmPunch.SubItems(I_LVW_FIN) = vbNullString
 lForeColor = COLOR_ROUGE
 End If

 If Not IsNull(rstPunch.Fields("NoClient")) And rstPunch.Fields("NoClient") <> vbNullString Then
 Call rstClient.Open("SELECT NomClient FROM GrbClient WHERE IDClient = " & rstPunch.Fields("NoClient"), g_connData, adOpenDynamic, adLockOptimistic)

 itmPunch.SubItems(I_LVW_CLIENT) = rstClient.Fields("NomClient")

 itmPunch.ListSubItems(I_LVW_CLIENT).Tag = rstPunch.Fields("NoClient")

1  Call rstClient.Close
 Else
 itmPunch.SubItems(I_LVW_CLIENT) = vbNullString
 End If

 If Not IsNull(rstPunch.Fields("Type")) Then
 If Left$(itmPunch.SubItems(I_LVW_PROJET), 1) = "E" Then
 itmPunch.SubItems(I_LVW_TYPE) = rstPunch.Fields("Type")
 
290
 Else
 itmPunch.SubItems(I_LVW_TYPE) = rstPunch.Fields("Type")
 End If
 End If

If Not IsNull(rstPunch.Fields("Commentaire")) And rstPunch.Fields("Commentaire") <> vbNullString Then
 itmPunch.SubItems(I_LVW_COMMENTAIRE) = rstPunch.Fields("Commentaire")
Else
 itmPunch.SubItems(I_LVW_COMMENTAIRE) = vbNullString
 End If

 If rstPunch.Fields("KM") = True Then
 If Not IsNull(rstPunch.Fields("NbreKM")) Then
4 itmPunch.SubItems(I_LVW_KM) = rstPunch.Fields("NbreKM")
4 Else
4 itmPunch.SubItems(I_LVW_KM) = 0
4 End If
4 Else
4 itmPunch.SubItems(I_LVW_KM) = ""
4 End If

4 lvwJour.ListItems(itmPunch.Index).ForeColor = lForeColor
4 lvwJour.ListItems(itmPunch.Index).ListSubItems(I_LVW_PROJET).ForeColor = lForeColor
4 lvwJour.ListItems(itmPunch.Index).ListSubItems(I_LVW_DEBUT).ForeColor = lForeColor
4 lvwJour.ListItems(itmPunch.Index).ListSubItems(I_LVW_FIN).ForeColor = lForeColor
4  lvwJour.ListItems(itmPunch.Index).ListSubItems(I_LVW_CLIENT).ForeColor = lForeColor
4  lvwJour.ListItems(itmPunch.Index).ListSubItems(I_LVW_TYPE).ForeColor = lForeColor
4  lvwJour.ListItems(itmPunch.Index).ListSubItems(I_LVW_COMMENTAIRE).ForeColor = lForeColor
4  lvwJour.ListItems(itmPunch.Index).ListSubItems(I_LVW_KM).ForeColor = lForeColor

4  Call rstPunch.MoveNext
4  Loop

4  Set rstEmploye = Nothing
4  Set rstClient = Nothing

50 Call rstPunch.Close
50 Set rstPunch = Nothing

 If lvwJour.ListItems.count > 0 Then
 lvwJour.ListItems(lvwJour.ListItems.count).Selected = True
 End If

 Exit Sub

Oups:

 wOups "frmPunch", "RemplirListViewJour", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListViewJourAutorisation()

 On Error GoTo Oups

 'Remplis ListView une journée
 Dim rstPunch As ADODB.Recordset
 Dim rstEmploye As ADODB.Recordset
 Dim rstAutorisation As ADODB.Recordset
 Dim rstClient As ADODB.Recordset
 Dim itmPunch As ListItem
 Dim lForeColor As Long

 Set rstAutorisation = New ADODB.Recordset

 Call rstAutorisation.Open("SELECT * FROM GrbAutorisationPunch WHERE AutoriserPar = " & m_iNoEmploye, g_connData, adOpenDynamic, adLockOptimistic)

 Set rstPunch = New ADODB.Recordset
 Set rstEmploye = New ADODB.Recordset
  Set rstClient = New ADODB.Recordset

  Do While Not rstAutorisation.EOF
  Call rstPunch.Open("SELECT * FROM GrbPunch WHERE Date = '" & ConvertDate(m_datDateChoisie) & "' AND NoEmploye = " & rstAutorisation.Fields("NoEmploye") & " ORDER BY HeureDébut", g_connData, adOpenDynamic, adLockOptimistic)

 'tant il y a de employé cedulé , ajoute dans lister
  Do While Not rstPunch.EOF
  Set itmPunch = lvwJour.ListItems.Add

  Call rstEmploye.Open("SELECT initiale FROM GrbEmployés WHERE NoEmploye = " & rstAutorisation.Fields("NoEmploye"), g_connData, adOpenDynamic, adLockOptimistic)

  itmPunch.Text = rstEmploye.Fields("initiale")

  Call rstEmploye.Close

 itmPunch.Tag = rstPunch.Fields("IDPunch")

If Not IsNull(rstPunch.Fields("NoProjet")) Then
 itmPunch.SubItems(I_LVW_PROJET) = rstPunch.Fields("NoProjet")
 Else
 itmPunch.SubItems(I_LVW_PROJET) = vbNullString
 End If

 If Not IsNull(rstPunch.Fields("HeureDébut")) Then
 itmPunch.SubItems(I_LVW_DEBUT) = rstPunch.Fields("HeureDébut")
 Else
 itmPunch.SubItems(I_LVW_DEBUT) = vbNullString
 End If

 If Not IsNull(rstPunch.Fields("HeureFin")) And rstPunch.Fields("HeureFin") <> vbNullString Then
 itmPunch.SubItems(I_LVW_FIN) = rstPunch.Fields("HeureFin")
 lForeColor = COLOR_NOIR
 Else
 itmPunch.SubItems(I_LVW_FIN) = vbNullString
 lForeColor = COLOR_ROUGE
 End If

 If Not IsNull(rstPunch.Fields("NoClient")) And rstPunch.Fields("NoClient") <> vbNullString Then
1  Call rstClient.Open("SELECT NomClient FROM GrbClient WHERE IDClient = " & rstPunch.Fields("NoClient"), g_connData, adOpenDynamic, adLockOptimistic)

 itmPunch.SubItems(I_LVW_CLIENT) = rstClient.Fields("NomClient")

 itmPunch.ListSubItems(I_LVW_CLIENT).Tag = rstPunch.Fields("NoClient")

 Call rstClient.Close
 Else
 itmPunch.SubItems(I_LVW_CLIENT) = vbNullString
 End If

 If Not IsNull(rstPunch.Fields("Type")) Then
 If Left$(itmPunch.SubItems(I_LVW_PROJET), 1) = "E" Then
 Select Case rstPunch.Fields("Type")
 Case "Dessin": itmPunch.SubItems(I_LVW_TYPE) = "Dessins électriques"
 Case "Fabrication": itmPunch.SubItems(I_LVW_TYPE) = "Fabrication"
 Case "Assemblage": itmPunch.SubItems(I_LVW_TYPE) = "Assemblage du panneau"
 Case "ProgInterface": itmPunch.SubItems(I_LVW_TYPE) = "Programmation d'interface"
 Case "ProgAutomate": itmPunch.SubItems(I_LVW_TYPE) = "Programmation d'automate"
 Case "ProgRobot": itmPunch.SubItems(I_LVW_TYPE) = "Programmation de robot"
 Case "Vision": itmPunch.SubItems(I_LVW_TYPE) = "Vision artificielle"
 Case "Test": itmPunch.SubItems(I_LVW_TYPE) = "Tests finaux"
 Case "Installation": itmPunch.SubItems(I_LVW_TYPE) = "Installation"
 Case "MiseService": itmPunch.SubItems(I_LVW_TYPE) = "Mise en service"
 Case "Formation": itmPunch.SubItems(I_LVW_TYPE) = "Formation du personnel"
 Case "Gestion": itmPunch.SubItems(I_LVW_TYPE) = "Gestion du projet"
 Case "Shipping": itmPunch.SubItems(I_LVW_TYPE) = "Expédition"
 Case "Prototypage-Dévelloppement expérimental": itmPunch.SubItems(I_LVW_TYPE) = "Prototypage-Dévelloppement expérimental"
 End Select
 Else
 Select Case rstPunch.Fields("Type")
 Case "Dessin": itmPunch.SubItems(I_LVW_TYPE) = "Conception et dessins"
 Case "Coupe": itmPunch.SubItems(I_LVW_TYPE) = "Coupe et préparation (sauf soudage)"
 Case "Machinage": itmPunch.SubItems(I_LVW_TYPE) = "Machinage"
 Case "Soudure": itmPunch.SubItems(I_LVW_TYPE) = "Coupe, soudure et meulage"
 Case "Assemblage": itmPunch.SubItems(I_LVW_TYPE) = "Assemblage des systèmes"
 Case "Peinture": itmPunch.SubItems(I_LVW_TYPE) = "Peinture et finition"
 Case "Test": itmPunch.SubItems(I_LVW_TYPE) = "Tests finaux"
 Case "Installation": itmPunch.SubItems(I_LVW_TYPE) = "Installation"
 Case "Formation": itmPunch.SubItems(I_LVW_TYPE) = "Formation du personnel"
 Case "Gestion": itmPunch.SubItems(I_LVW_TYPE) = "Gestion du projet"
 Case "Shipping": itmPunch.SubItems(I_LVW_TYPE) = "Expédition"
 Case "Prototypage-Dévelloppement expérimental": itmPunch.SubItems(I_LVW_TYPE) = "Prototypage-Dévelloppement expérimental"
 End Select
 End If
 End If

 If Not IsNull(rstPunch.Fields("Commentaire")) And rstPunch.Fields("Commentaire") <> vbNullString Then
 itmPunch.SubItems(I_LVW_COMMENTAIRE) = rstPunch.Fields("Commentaire")
 Else
 itmPunch.SubItems(I_LVW_COMMENTAIRE) = vbNullString
4 End If

4 If rstPunch.Fields("KM") = True Then
4 itmPunch.SubItems(I_LVW_KM) = rstPunch.Fields("NbreKM")
4 Else
4 itmPunch.SubItems(I_LVW_KM) = vbNullString
4 End If

4 lvwJour.ListItems(itmPunch.Index).ForeColor = lForeColor
4 lvwJour.ListItems(itmPunch.Index).ListSubItems(I_LVW_PROJET).ForeColor = lForeColor
4 lvwJour.ListItems(itmPunch.Index).ListSubItems(I_LVW_DEBUT).ForeColor = lForeColor
4 lvwJour.ListItems(itmPunch.Index).ListSubItems(I_LVW_FIN).ForeColor = lForeColor
4 lvwJour.ListItems(itmPunch.Index).ListSubItems(I_LVW_CLIENT).ForeColor = lForeColor
4  lvwJour.ListItems(itmPunch.Index).ListSubItems(I_LVW_TYPE).ForeColor = lForeColor
4  lvwJour.ListItems(itmPunch.Index).ListSubItems(I_LVW_COMMENTAIRE).ForeColor = lForeColor
4  lvwJour.ListItems(itmPunch.Index).ListSubItems(I_LVW_KM).ForeColor = lForeColor

4  Call rstPunch.MoveNext
4  Loop

4  Call rstPunch.Close

4  Call rstAutorisation.MoveNext
4  Loop

50 Set rstPunch = Nothing
50 Set rstClient = Nothing
 Set rstEmploye = Nothing

 Call rstAutorisation.Close
 Set rstAutorisation = Nothing

 Call lvwJour_Click

 Exit Sub

Oups:

 wOups "frmPunch", "RemplirListViewJourAutorisation", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListViewSemaine(ByVal bAujourdhui As Boolean)

 On Error GoTo Oups

 'remplis une semaine
 'bAujourdhui sert à savoir si on rafraichit seulement la journée d'aujourd'hui
 Dim rstPunch As ADODB.Recordset
 Dim rstEmploye As ADODB.Recordset
 Dim iJourSemaine As Integer
 Dim datPremiereDate As Date
 Dim datDerniereDate As Date
 Dim iCompteur As Integer
 Dim sHeureDebutFin As String
 Dim itmSemaine As ListItem
 Dim lForeColor As Long

 Set rstPunch = New ADODB.Recordset
  Set rstEmploye = New ADODB.Recordset

  If bAujourdhui = False Then
  For iCompteur = 1 To 7
 'couleur par defaut entete de date
  lbljour(iCompteur - 1).ForeColor = vbWhite
  lblNomJour(iCompteur - 1).ForeColor = vbWhite

  Call lvwJourSemaine(iCompteur).ListItems.Clear
  Next

  iJourSemaine = Weekday(m_datDateChoisie)
datPremiereDate = m_datDateChoisie
1 datDerniereDate = m_datDateChoisie

 'Trouve premiere date de la semaine
 Do While Not Weekday(datPremiereDate) = 1
 datPremiereDate = datPremiereDate - 1
 Loop

 'Trouve derniere date de la semaine
 Do While Not Weekday(datDerniereDate) = 7
 datDerniereDate = datDerniereDate + 1
 Loop

 'Sélectionne la semaine courante
 Call rstPunch.Open("SELECT * FROM GrbPunch WHERE Date BETWEEN '" & ConvertDate(datPremiereDate) & "' AND '" & ConvertDate(datDerniereDate) & "' AND NoEmploye = " & m_iNoEmploye & " ORDER BY HeureDébut", g_connData, adOpenDynamic, adLockOptimistic)

 For iCompteur = 1 To 7
 'Pour écrire le jour
 lbljour(iCompteur - 1).Caption = Day(datPremiereDate + iCompteur - 1)

 'Garde en memoire la date des listers
 lvwJourSemaine(iCompteur).Tag = datPremiereDate + iCompteur - 1
Next

 Do While Not rstPunch.EOF
 'ajoute dans le lister, dépendant le jour de la semaine
 Set itmSemaine = lvwJourSemaine(Weekday(rstPunch.Fields("Date"))).ListItems.Add

 itmSemaine.Tag = rstPunch.Fields("IDPunch")

 Call rstEmploye.Open("SELECT initiale FROM GrbEmployés WHERE NoEmploye = " & m_iNoEmploye, g_connData, adOpenDynamic, adLockOptimistic)

 itmSemaine.Text = rstEmploye.Fields("initiale")

 Call rstEmploye.Close

1  sHeureDebutFin = Trim(rstPunch.Fields("HeureDébut") + "-")

 If Not IsNull(rstPunch.Fields("HeureFin")) And rstPunch.Fields("HeureFin") <> vbNullString Then
 sHeureDebutFin = sHeureDebutFin + rstPunch.Fields("HeureFin")
 lForeColor = COLOR_NOIR
 Else
 lForeColor = COLOR_ROUGE
 End If

 itmSemaine.SubItems(I_LVW_TEMPS) = sHeureDebutFin

 lvwJourSemaine(Weekday(rstPunch.Fields("Date"))).ListItems(itmSemaine.Index).ForeColor = lForeColor
 lvwJourSemaine(Weekday(rstPunch.Fields("Date"))).ListItems(itmSemaine.Index).ListSubItems(I_LVW_TEMPS).ForeColor = lForeColor

 Call rstPunch.MoveNext
 Loop

 Call rstPunch.Close
2  Else
 Call lvwJourSemaine(Weekday(m_datDateChoisie)).ListItems.Clear

Call rstPunch.Open("SELECT * FROM GrbPunch WHERE Date = '" & ConvertDate(m_datDateChoisie) & "' AND NoEmploye = " & m_iNoEmploye, g_connData, adOpenDynamic, adLockOptimistic)

 Do While Not rstPunch.EOF
 'ajoute dans le lister, dépendant le jour de la semaine
 Set itmSemaine = lvwJourSemaine(Weekday(rstPunch.Fields("Date"))).ListItems.Add

 itmSemaine.Tag = rstPunch.Fields("IDPunch")

 Call rstEmploye.Open("SELECT initiale FROM GrbEmployés WHERE NoEmploye = " & m_iNoEmploye, g_connData, adOpenDynamic, adLockOptimistic)

 itmSemaine.Text = rstEmploye.Fields("initiale")

 Call rstEmploye.Close

sHeureDebutFin = Trim(rstPunch.Fields("HeureDébut") + "-")

 If Not IsNull(rstPunch.Fields("HeureFin")) And rstPunch.Fields("HeureFin") <> vbNullString Then
 sHeureDebutFin = sHeureDebutFin + rstPunch.Fields("HeureFin")
 lForeColor = COLOR_NOIR
 Else
 lForeColor = COLOR_ROUGE
 End If

 itmSemaine.SubItems(I_LVW_TEMPS) = sHeureDebutFin

 lvwJourSemaine(Weekday(m_datDateChoisie)).ListItems(itmSemaine.Index).ForeColor = lForeColor
 lvwJourSemaine(Weekday(m_datDateChoisie)).ListItems(itmSemaine.Index).ListSubItems(I_LVW_TEMPS).ForeColor = lForeColor

 Call rstPunch.MoveNext
Loop

 Call rstPunch.Close
3  End If

Set rstPunch = Nothing
3  Set rstEmploye = Nothing

Exit Sub

Oups:

3  wOups "frmPunch", "RemplirListViewSemaine", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListViewSemaineAutorisation(ByVal bAujourdhui As Boolean)

 On Error GoTo Oups

 'remplis une semaine
 Dim rstPunch As ADODB.Recordset
 Dim rstEmploye As ADODB.Recordset
 Dim rstAutorisation As ADODB.Recordset
 Dim iJourSemaine As Integer
 Dim datPremiereDate As Date
 Dim datDerniereDate As Date
 Dim iCompteur As Integer
 Dim sHeureDebutFin As String
 Dim itmSemaine As ListItem
 Dim lForeColor As Long

  Set rstPunch = New ADODB.Recordset
  Set rstEmploye = New ADODB.Recordset
  Set rstAutorisation = New ADODB.Recordset

  If bAujourdhui = False Then
  For iCompteur = 1 To 7
 'couleur par defaut entete de date
  lbljour(iCompteur - 1).ForeColor = vbWhite
  lblNomJour(iCompteur - 1).ForeColor = vbWhite
  Next
 
iJourSemaine = Weekday(m_datDateChoisie)
1 datPremiereDate = m_datDateChoisie
 datDerniereDate = m_datDateChoisie
 
 'trouve premiere date de la semaine
 Do While Not Weekday(datPremiereDate) = 1
 datPremiereDate = datPremiereDate - 1
 Loop
 
 'trouve derniere date de la semaine
 Do While Not Weekday(datDerniereDate) = 7
 datDerniereDate = datDerniereDate + 1
 Loop
 
 For iCompteur = 1 To 7
 'pour ecrire le jour
 lbljour(iCompteur - 1).Caption = Day(datPremiereDate + iCompteur - 1)
 
 'garde en memoire la date des lister
 lvwJourSemaine(iCompteur).Tag = datPremiereDate + iCompteur - 1
Next
 
 Call rstAutorisation.Open("SELECT * FROM GrbAutorisationPunch WHERE AutoriserPar = " & m_iNoEmploye, g_connData, adOpenDynamic, adLockOptimistic)
 
 Do While Not rstAutorisation.EOF
 'selectionne la semaine courante
 Call rstPunch.Open("SELECT * FROM GrbPunch WHERE Date BETWEEN '" & ConvertDate(datPremiereDate) & "' AND '" & ConvertDate(datDerniereDate) & "' AND NoEmploye = " & rstAutorisation.Fields("NoEmploye") & " ORDER BY HeureDébut", g_connData, adOpenDynamic, adLockOptimistic)
 
 Do While Not rstPunch.EOF
 'ajoute dans le lister, dépendant le jour de la semaine
 Set itmSemaine = lvwJourSemaine(Weekday(rstPunch.Fields("Date"))).ListItems.Add
 
 itmSemaine.Tag = rstPunch.Fields("IDPunch")
 
1  Call rstEmploye.Open("SELECT initiale FROM GrbEmployés WHERE NoEmploye = " & rstAutorisation.Fields("NoEmploye"), g_connData, adOpenDynamic, adLockOptimistic)
 
 itmSemaine.Text = rstEmploye.Fields("initiale")
 
 Call rstEmploye.Close
 
 sHeureDebutFin = Trim(rstPunch.Fields("HeureDébut") + "-")
 
 If Not IsNull(rstPunch.Fields("HeureFin")) And rstPunch.Fields("HeureFin") <> vbNullString Then
 sHeureDebutFin = sHeureDebutFin + rstPunch.Fields("HeureFin")
 lForeColor = COLOR_NOIR
 Else
 lForeColor = COLOR_ROUGE
 End If
 
 itmSemaine.SubItems(I_LVW_TEMPS) = sHeureDebutFin
 
 lvwJourSemaine(Weekday(rstPunch.Fields("Date"))).ListItems(itmSemaine.Index).ForeColor = lForeColor
 lvwJourSemaine(Weekday(rstPunch.Fields("Date"))).ListItems(itmSemaine.Index).ListSubItems(I_LVW_TEMPS).ForeColor = lForeColor
 
 Call rstPunch.MoveNext
 Loop
 
 Call rstPunch.Close
 
 Call rstAutorisation.MoveNext
Loop
 
 Call rstAutorisation.Close
Else
Call rstAutorisation.Open("SELECT * FROM GrbAutorisationPunch WHERE AutoriserPar = " & m_iNoEmploye, g_connData, adOpenDynamic, adLockOptimistic)
 
3 Do While Not rstAutorisation.EOF
 'selectionne la semaine courante
 Call rstPunch.Open("SELECT * FROM GrbPunch WHERE Date = '" & ConvertDate(m_datDateChoisie) & "' AND NoEmploye = " & rstAutorisation.Fields("NoEmploye") & " ORDER BY HeureDébut", g_connData, adOpenDynamic, adLockOptimistic)
 
 Do While Not rstPunch.EOF
 'ajoute dans le lister, dépendant le jour de la semaine
 Set itmSemaine = lvwJourSemaine(Weekday(m_datDateChoisie)).ListItems.Add
 
 itmSemaine.Tag = rstPunch.Fields("IDPunch")
 
 Call rstEmploye.Open("SELECT initiale FROM GrbEmployés WHERE NoEmploye = " & rstAutorisation.Fields("NoEmploye"), g_connData, adOpenDynamic, adLockOptimistic)
 
 itmSemaine.Text = rstEmploye.Fields("initiale")
 
 Call rstEmploye.Close
 
 sHeureDebutFin = Trim(rstPunch.Fields("HeureDébut") + "-")
 
 If Not IsNull(rstPunch.Fields("HeureFin")) And rstPunch.Fields("HeureFin") <> vbNullString Then
 sHeureDebutFin = sHeureDebutFin + rstPunch.Fields("HeureFin")
 lForeColor = COLOR_NOIR
 Else
 lForeColor = COLOR_ROUGE
 End If
 
 itmSemaine.SubItems(I_LVW_TEMPS) = sHeureDebutFin
 
 lvwJourSemaine(Weekday(m_datDateChoisie)).ListItems(itmSemaine.Index).ForeColor = lForeColor
 lvwJourSemaine(Weekday(m_datDateChoisie)).ListItems(itmSemaine.Index).ListSubItems(I_LVW_TEMPS).ForeColor = lForeColor
 
 Call rstPunch.MoveNext
 Loop
 
4 Call rstPunch.Close
 
4 Call rstAutorisation.MoveNext
4 Loop
 
4 Call rstAutorisation.Close
4 End If

4 Set rstAutorisation = Nothing
4 Set rstEmploye = Nothing
4 Set rstPunch = Nothing

4 Exit Sub

Oups:

4 wOups "frmPunch", "RemplirListViewSemaineAutorisation", Err, Err.number, Err.Description
End Sub

Private Sub chkDiner_Click()

 On Error GoTo Oups

 If chkDiner.Value = vbChecked Then
 optHeureDiner(I_OPT_1_HEURE).Enabled = True
 optHeureDiner(I_OPT_30_MINUTES).Enabled = True
 Else
 optHeureDiner(I_OPT_1_HEURE).Enabled = False
 optHeureDiner(I_OPT_30_MINUTES).Enabled = False
 End If

 m_bMonthViewHasFocus = False

 Exit Sub

Oups:

 wOups "frmPunch", "chkDiner_Click", Err, Err.number, Err.Description
End Sub

Private Sub chkPMDiner_Click()

 On Error GoTo Oups

 If chkPMDiner.Value = vbChecked Then
 optPMHeureDiner(I_OPT_1_HEURE).Enabled = True
 optPMHeureDiner(I_OPT_30_MINUTES).Enabled = True
 Else
 optPMHeureDiner(I_OPT_1_HEURE).Enabled = False
 optPMHeureDiner(I_OPT_30_MINUTES).Enabled = False
 End If

 m_bMonthViewHasFocus = False

 Exit Sub

Oups:

  wOups "frmPunch", "chkPMDiner_Click", Err, Err.number, Err.Description
End Sub

Private Sub chkDiner_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 On Error GoTo Oups

 If m_bMonthViewHasFocus = True Then
 If chkDiner.Value = vbChecked Then
 chkDiner.Value = vbUnchecked
 Else
 chkDiner.Value = vbChecked
 End If
 End If

 Exit Sub

Oups:

 wOups "frmPunch", "chkDiner_MouseUp", Err, Err.number, Err.Description
End Sub

Private Sub chkPMDiner_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 On Error GoTo Oups

 If m_bMonthViewHasFocus = True Then
 If chkPMDiner.Value = vbChecked Then
 chkPMDiner.Value = vbUnchecked
 Else
 chkPMDiner.Value = vbChecked
 End If
 End If

 Exit Sub

Oups:

 wOups "frmPunch", "chkPMDiner_MouseUp", Err, Err.number, Err.Description
End Sub

Private Sub chkKM_Click()

 On Error GoTo Oups

 If chkKM.Value = vbChecked Then
 txtKM.Enabled = True
 Else
 txtKM.Text = ""
 txtKM.Enabled = False
 End If

 m_bMonthViewHasFocus = False

 Exit Sub

Oups:

 wOups "frmPunch", "chkKM_Click", Err, Err.number, Err.Description
End Sub

Private Sub chkKM_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 On Error GoTo Oups

 If m_bMonthViewHasFocus = True Then
 If chkKM.Value = vbChecked Then
 chkKM.Value = vbUnchecked
 Else
 chkKM.Value = vbChecked
 End If
 End If

 Exit Sub

Oups:

 wOups "frmPunch", "chkKM_MouseUp", Err, Err.number, Err.Description
End Sub

Private Sub cmbEmployé_Click()

 On Error GoTo Oups

 txtEmploye.Text = cmbemployé.Text

 Exit Sub

Oups:

 wOups "frmPunch", "cmbEmployé_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmbHeureSemaine_Click()
 
 On Error GoTo Oups

 Call CalculerHeureSemaine

 Exit Sub

Oups:

 wOups "frmPunch", "cmbHeureSemaine_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnuler_Click()

 On Error GoTo Oups

 frajour.Visible = True
 fraPunch.Visible = False

 m_bMonthViewHasFocus = False

 Exit Sub

Oups:

 wOups "frmPunch", "cmdAnnuler_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdBrowseComment_Click()

 On Error GoTo Oups

 Dim sProjet As String

 If mskNoProjet.Visible = True Then
 sProjet = mskNoProjet.Text
 Else
 sProjet = txtnoprojet.Text
 End If

 If txtnoprojet.Text <> "" Or mskNoProjet.Text <> "" Then
 If txtClient.Text <> "" Then
 Call frmChoixCommentaire.Afficher(sProjet)

 If sCommentaire <> "" Then
  txtCommentaires.Text = sCommentaire
  End If
  Else
  Call MsgBox("Numéro de projet ou soumission invalide!", vbOKOnly, "Erreur")
  End If
  Else
  Call MsgBox("Le numéro de projet ou soumission est obligatoire!", vbOKOnly, "Erreur")
  End If

10 Exit Sub

Oups:

wOups "frmPunch", "cmdBrowseComment_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdBrowseCommentPM_Click()

 On Error GoTo Oups

 If mskPMNoProjet.Text <> "" Then
 If txtPMClient.Text <> "" Then
 Call frmChoixCommentaire.Afficher(mskPMNoProjet.Text)

 txtCommentaires.Text = sCommentaire
 Else
 Call MsgBox("Numéro de projet ou soumission invalide!", vbOKOnly, "Erreur")
 End If
 Else
 Call MsgBox("Le numéro de projet ou soumission est obligaoire!", vbOKOnly, "Erreur")
 End If

  Exit Sub

Oups:

  wOups "frmPunch", "cmdBrowseCommentPM_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdPMAnnuler_Click()

 On Error GoTo Oups

 frajour.Visible = True
 fraPunch.Visible = False
 fraPunchMultiple.Visible = False

 m_bMonthViewHasFocus = False

 Exit Sub

Oups:

 wOups "frmPunch", "cmdPMAnnuler_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnuler_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 On Error GoTo Oups

 If m_bMonthViewHasFocus = True Then
 Call cmdAnnuler_Click
 End If

 Exit Sub

Oups:

 wOups "frmPunch", "cmdAnnuler_MouseUp", Err, Err.number, Err.Description
End Sub

Private Sub cmdPMAnnuler_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 On Error GoTo Oups

 If m_bMonthViewHasFocus = True Then
 Call cmdPMAnnuler_Click
 End If

 Exit Sub

Oups:

 wOups "frmPunch", "cmdPMAnnuler_MouseUp", Err, Err.number, Err.Description
End Sub

Private Sub cmdModifierPunchIn_Click()

 On Error GoTo Oups

 Dim rstEmploye As ADODB.Recordset
 Dim itmPunch As ListItem
 Dim iCompteur As Integer

 If VerifierModificationDate = True Then
 Set itmPunch = lvwJour.SelectedItem
 
 Set rstEmploye = New ADODB.Recordset
 
 Call rstEmploye.Open("SELECT employe FROM GrbEmployés WHERE Initiale = '" & itmPunch.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 For iCompteur = 0 To cmbemployé.ListCount - 1
 If cmbemployé.LIST(iCompteur) = rstEmploye.Fields("employe") Then
 cmbemployé.ListIndex = iCompteur
 
  Exit For
  End If
  Next
 
  Call rstEmploye.Close
  Set rstEmploye = Nothing
 
  txtClient.Text = itmPunch.SubItems(I_LVW_CLIENT)
 
  txtClient.Tag = itmPunch.ListSubItems(I_LVW_CLIENT).Tag
 
  cmbemployé.Visible = True
txtEmploye.Visible = False
 
1 Select Case Left$(itmPunch.SubItems(I_LVW_PROJET), 1)
 Case "E": optTypePunch(I_OPT_ELECTRIQUE).Value = True
 Case "M": optTypePunch(I_OPT_MECANIQUE).Value = True
 End Select
 
 mskNoProjet.Text = Right$(itmPunch.SubItems(I_LVW_PROJET), 8)
 
 m_ePunch = I_MODIF_PUNCH_IN
 
 Call RemplirComboType
 
 If Left$(itmPunch.SubItems(I_LVW_PROJET), 1) = "E" Then
 
 If Not IsNull(itmPunch.SubItems(I_LVW_TYPE)) Then
 cmbType.Text = itmPunch.SubItems(I_LVW_TYPE)
 Else
 
 cmbType.ListIndex = -1
 End If

 Else
 If Not itmPunch.SubItems(I_LVW_TYPE) = "Soumission" Then
 If Not IsNull(itmPunch.SubItems(I_LVW_TYPE)) Then
 cmbType.Text = itmPunch.SubItems(I_LVW_TYPE)
 
 Else
 cmbType.ListIndex = -1
 End If
 End If
 End If
 
mskNoProjet.Visible = True
 txtnoprojet.Visible = False
 
picTypePunch.Enabled = True
 
3 mskHeure.mask = "##:##"
 mskHeure.Text = itmPunch.SubItems(I_LVW_DEBUT)
 
 m_bModifPunch = True
 
 optHeure(I_OPT_MANUELLEMENT).Value = True
 
 m_bModifPunch = False
 
 txtCommentaires.Text = itmPunch.SubItems(I_LVW_COMMENTAIRE)
 
 If itmPunch.SubItems(I_LVW_KM) <> "" Then
 chkKM.Value = vbChecked
 txtKM.Text = itmPunch.SubItems(I_LVW_KM)
 Else
 chkKM.Value = vbUnchecked
 txtKM.Text = vbNullString
 End If
 
fraPunch.Caption = "Modification du punch in"
 
 frajour.Visible = False
fraPunchMultiple.Visible = False
 fraPunch.Visible = True
 
 chkDiner.Visible = False
 optHeureDiner(I_OPT_30_MINUTES).Visible = False
optHeureDiner(I_OPT_1_HEURE).Visible = False
End If

4 m_bMonthViewHasFocus = False

4 Exit Sub

Oups:

4 wOups "frmPunch", "cmdModifierPunchIn_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdModifierPunchIn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 On Error GoTo Oups

 If m_bMonthViewHasFocus = True Then
 Call cmdModifierPunchIn_Click
 End If

 Exit Sub

Oups:

 wOups "frmPunch", "cmdModifierPunchIn_MouseUp", Err, Err.number, Err.Description
End Sub

Private Sub cmdModifierPunchOut_Click()

 On Error GoTo Oups

 Dim rstEmploye As ADODB.Recordset

 If VerifierModificationDate = True Then
 If lvwJour.ListItems.count > 0 Then
 m_ePunch = I_MODIF_PUNCH_OUT
 
 Call AfficherPunchOut
 
 fraPunch.Caption = "Modification du punch out"
 
 chkDiner.Visible = True
 optHeureDiner(I_OPT_30_MINUTES).Visible = True
 optHeureDiner(I_OPT_1_HEURE).Visible = True
 
 Set rstEmploye = New ADODB.Recordset
 
  Call rstEmploye.Open("SELECT GrbFamille.Famille FROM Grbemployés INNER JOIN GrbFamille ON Grbemployés.Famille = GrbFamille.IDFamille WHERE employe = '" & m_iNoEmploye & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
  If Not rstEmploye.EOF Then
  Select Case rstEmploye.Fields("Famille")
 Case "Administration": optHeureDiner(I_OPT_1_HEURE).Value = True
  Case "Technicien": optHeureDiner(I_OPT_1_HEURE).Value = True
  Case Else: optHeureDiner(I_OPT_30_MINUTES).Value = True
  End Select
  Else
  optHeureDiner(I_OPT_30_MINUTES).Value = True
 End If
 
Call rstEmploye.Close
 Set rstEmploye = Nothing
 
 chkDiner.Value = vbUnchecked
 
 m_bModifPunch = True
 
 optHeure(I_OPT_MANUELLEMENT).Value = True
 
 m_bModifPunch = False
 
 mskHeure.mask = "##:##"
 mskHeure.Text = lvwJour.SelectedItem.SubItems(I_LVW_FIN)
 End If
End If

m_bMonthViewHasFocus = False

1  Exit Sub

Oups:

wOups "frmPunch", "cmdModifierPunchOut_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdModifierPunchOut_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 On Error GoTo Oups

 If m_bMonthViewHasFocus = True Then
 Call cmdModifierPunchOut_Click
 End If

 Exit Sub

Oups:

 wOups "frmPunch", "cmdModifierPunchOut_MouseUp", Err, Err.number, Err.Description
End Sub

Private Sub cmdOK_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
 On Error GoTo Oups

 If m_bMonthViewHasFocus = True Then
 Call cmdOK_Click
 End If

 Exit Sub
 
Oups:

 wOups "frmPunch", "cmdOK_MouseUp", Err, Err.number, Err.Description
End Sub

Private Sub cmdPMOK_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
 On Error GoTo Oups

 If m_bMonthViewHasFocus = True Then
 Call cmdPMOK_Click
 End If

 Exit Sub
 
Oups:

 wOups "frmPunch", "cmdPMOK_MouseUp", Err, Err.number, Err.Description
End Sub

Private Sub cmdPunchIn_Click()

 On Error GoTo Oups
 
 Dim rstEmploye As ADODB.Recordset
 Dim iCompteur As Integer
 
 If VerifierModificationDate = True Then
 mskNoProjet.mask = vbNullString
 mskNoProjet.Text = vbNullString
 mskNoProjet.mask = "#####-##"
 
 txtClient.Text = vbNullString
 
 cmbemployé.Visible = True
 txtEmploye.Visible = False
 
 mskNoProjet.Visible = True
  txtnoprojet.Visible = False
 
  picTypePunch.Enabled = True
 
  Set rstEmploye = New ADODB.Recordset
 
  Call rstEmploye.Open("SELECT GrbFamille.Famille FROM Grbemployés INNER JOIN GrbFamille ON Grbemployés.Famille = GrbFamille.IDFamille WHERE employe = '" & m_iNoEmploye & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
  If Not rstEmploye.EOF Then
  Select Case rstEmploye.Fields("Famille")
 Case "Électrique": optTypePunch(I_OPT_ELECTRIQUE).Value = True
  Case "Mécanique": optTypePunch(I_OPT_MECANIQUE).Value = True
  Case Else: optTypePunch(I_OPT_ELECTRIQUE).Value = True
 End Select
1 Else
 optTypePunch(I_OPT_ELECTRIQUE).Value = True
 End If
 
 Call rstEmploye.Close
 Set rstEmploye = Nothing
 
 cmbType.ListIndex = -1
 
 
 mskHeure.mask = vbNullString
 mskHeure.Text = vbNullString
 mskHeure.mask = "##:##"
 
 optHeure(I_OPT_SYSTEME).Value = True
 
 txtCommentaires.Text = vbNullString
 
chkKM.Value = vbUnchecked
 
 txtKM.Text = vbNullString
 
 fraPunch.Caption = "Punch in"
 
 frajour.Visible = False
 fraPunch.Visible = True
 fraPunchMultiple.Visible = False
 
 chkDiner.Visible = False
1  optHeureDiner(I_OPT_30_MINUTES).Visible = False
 optHeureDiner(I_OPT_1_HEURE).Visible = False
 
 Call mskNoProjet.SetFocus
 
 m_ePunch = I_PUNCH_IN
End If

m_bMonthViewHasFocus = False

Exit Sub

Oups:

wOups "frmPunch", "cmdPunchIn_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdPunchIn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 On Error GoTo Oups

 If m_bMonthViewHasFocus = True Then
 Call cmdPunchIn_Click
 End If

 Exit Sub

Oups:

 wOups "frmPunch", "cmdPunchIn_MouseUp", Err, Err.number, Err.Description
End Sub

Private Sub cmdPunchMultiple_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 On Error GoTo Oups

 If m_bMonthViewHasFocus = True Then
 Call cmdPunchMultiple_Click
 End If

 Exit Sub

Oups:

 wOups "frmPunch", "cmdPunchMultiple_MouseUp", Err, Err.number, Err.Description
End Sub

Private Sub cmdPunchMultiple_Click()

 On Error GoTo Oups

 Dim iCompteur As Integer
 Dim rstEmploye As ADODB.Recordset

 If VerifierModificationDate = True Then
 For iCompteur = 1 To lvwEmployes.ListItems.count
 lvwEmployes.ListItems(iCompteur).Checked = False
 Next
 
 mskPMNoProjet.mask = vbNullString
 mskPMNoProjet.Text = vbNullString
 mskPMNoProjet.mask = "#####-##"
 
 Set rstEmploye = New ADODB.Recordset

  Call rstEmploye.Open("SELECT GrbFamille.Famille FROM Grbemployés INNER JOIN GrbFamille ON Grbemployés.Famille = GrbFamille.IDFamille WHERE employe = '" & m_iNoEmploye & "'", g_connData, adOpenDynamic, adLockOptimistic)

  If Not rstEmploye.EOF Then
  Select Case rstEmploye.Fields("Famille")
 Case "Électrique": optPMTypePunch(I_OPT_ELECTRIQUE).Value = True
  Case "Mécanique": optPMTypePunch(I_OPT_MECANIQUE).Value = True
  Case Else: optPMTypePunch(I_OPT_ELECTRIQUE).Value = True
  End Select
  Else
  optPMTypePunch(I_OPT_ELECTRIQUE).Value = True
End If

1 Call rstEmploye.Close
 Set rstEmploye = Nothing
 
 mskPMHeureDebut.mask = vbNullString
 mskPMHeureDebut.Text = vbNullString
 mskPMHeureDebut.mask = "##:##"

 mskPMHeureFin.mask = vbNullString
 mskPMHeureFin.Text = vbNullString
 mskPMHeureFin.mask = "##:##"
 
 txtPMCommentaire.Text = vbNullString
 
 chkPMDiner.Value = vbUnchecked

 fraPunch.Visible = False
frajour.Visible = False
 fraPunchMultiple.Visible = True
 End If

m_bMonthViewHasFocus = False

 Exit Sub

Oups:

wOups "frmPunch", "cmdPunchMultiple_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdPunchOut_Click()

 On Error GoTo Oups

 Dim rstEmploye As ADODB.Recordset

 If VerifierModificationDate = True Then
 If lvwJour.ListItems.count > 0 Then
 If lvwJour.SelectedItem.ForeColor = COLOR_ROUGE Then
 m_ePunch = I_PUNCH_OUT
 
 Call AfficherPunchOut
 
 fraPunch.Caption = "Punch out"
 
 chkDiner.Visible = True
 optHeureDiner(I_OPT_30_MINUTES).Visible = True
 optHeureDiner(I_OPT_1_HEURE).Visible = True
 
  Set rstEmploye = New ADODB.Recordset
 
  Call rstEmploye.Open("SELECT GrbFamille.Famille FROM Grbemployés INNER JOIN GrbFamille ON Grbemployés.Famille = GrbFamille.IDFamille WHERE employe = '" & m_iNoEmploye & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
  If Not rstEmploye.EOF Then
  Select Case rstEmploye.Fields("Famille")
 Case "Administration": optHeureDiner(I_OPT_1_HEURE).Value = True
  Case "Technicien": optHeureDiner(I_OPT_1_HEURE).Value = True
  Case Else: optHeureDiner(I_OPT_30_MINUTES).Value = True
  End Select
  Else
 optHeureDiner(I_OPT_30_MINUTES).Value = True
 End If
 
 Call rstEmploye.Close
 Set rstEmploye = Nothing
 
 chkDiner.Value = vbUnchecked
 
 mskHeure.mask = vbNullString
 mskHeure.Text = vbNullString
 mskHeure.mask = "##:##"
 
 optHeure(I_OPT_SYSTEME).Value = True
 Else
 Call MsgBox("Le punch out a déjà été fait!")
 End If
End If
End If

 m_bMonthViewHasFocus = False

Exit Sub

Oups:

 wOups "frmPunch", "cmdPunchOut_Click", Err, Err.number, Err.Description
End Sub

Private Sub AfficherPunchOut()

 On Error GoTo Oups

 Dim rstEmploye As ADODB.Recordset
 Dim rstPunch As ADODB.Recordset
 Dim rstClient As ADODB.Recordset
 Dim iCompteur As Integer
 Dim G As Integer
 Set rstPunch = New ADODB.Recordset
 Set rstEmploye = New ADODB.Recordset
 Set rstClient = New ADODB.Recordset
 
 Call rstPunch.Open("SELECT * FROM GrbPunch WHERE IDPunch = " & lvwJour.SelectedItem.Tag, g_connData, adOpenDynamic, adLockOptimistic)
 
 Call rstEmploye.Open("SELECT employe FROM GrbEmployés WHERE NoEmploye = " & rstPunch.Fields("NoEmploye"), g_connData, adOpenDynamic, adLockOptimistic)
 
 For iCompteur = 0 To cmbemployé.ListCount - 1
  If cmbemployé.LIST(iCompteur) = rstEmploye.Fields("Employe") Then
  cmbemployé.ListIndex = iCompteur

  Exit For
  End If
  Next
 
  Call rstEmploye.Close
  Set rstEmploye = Nothing
 
  Call rstClient.Open("SELECT NomClient FROM GrbClient WHERE IDClient = " & rstPunch.Fields("NoClient"), g_connData, adOpenDynamic, adLockOptimistic)
 
10 txtClient.Text = rstClient.Fields("NomClient")

txtClient.Tag = rstPunch.Fields("NoClient")
 
Call rstClient.Close
Set rstClient = Nothing
 
txtnoprojet.Text = Right(rstPunch.Fields("NoProjet"), 8)
 
Call RemplirComboType
 
Call AfficherTypePunch
 
If Not IsNull(rstPunch.Fields("Commentaire")) Then
 txtCommentaires.Text = rstPunch.Fields("Commentaire")
Else
 txtCommentaires.Text = vbNullString
End If

1  If rstPunch.Fields("KM") = True Then
 chkKM.Value = vbChecked

 If Not IsNull(rstPunch.Fields("NbreKM")) Then
 txtKM.Text = rstPunch.Fields("NbreKM")
 Else
 txtKM.Text = 0
 End If
1  Else
 chkKM.Value = vbUnchecked
 txtKM.Text = vbNullString
End If
 
Select Case Left(rstPunch.Fields("NoProjet"), 1)
 Case "E": optTypePunch(I_OPT_ELECTRIQUE).Value = True
 Case "M": optTypePunch(I_OPT_MECANIQUE).Value = True
End Select

If Not IsNull(rstPunch.Fields("Type")) Then
 If Left(rstPunch.Fields("NoProjet"), 1) = "E" Then
 For G = 0 To cmbType.ListCount
 If cmbType.LIST(G) = rstPunch.Fields("Type") Then
 cmbType.ListIndex = G
 Exit For
 End If
 Next
310
 Else
 For G = 0 To cmbType.ListCount
 If cmbType.LIST(G) = rstPunch.Fields("Type") Then
 cmbType.ListIndex = G
 Exit For
 End If
 Next
 End If
3  End If
 
 picTypePunch.Enabled = False

40 txtnoprojet.Visible = True
mskNoProjet.Visible = False
 
4 txtEmploye.Visible = True
4 cmbemployé.Visible = False
 
4 frajour.Visible = False
4 fraPunch.Visible = True
4 fraPunchMultiple.Visible = False

4 Exit Sub

Oups:

4 wOups "frmPunch", "AfficherPunchOut", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboEmploye()

 On Error GoTo Oups

 Dim rstEmploye As ADODB.Recordset
 Dim rstAutorisation As ADODB.Recordset
 Dim itmEmploye As ListItem
 
 Call cmbemployé.Clear
 Call lvwEmployes.ListItems.Clear
 
 Set rstEmploye = New ADODB.Recordset
 Set rstAutorisation = New ADODB.Recordset
 
 Call rstEmploye.Open("SELECT Employe FROM GrbEmployés WHERE NoEmploye = " & m_iNoEmploye, g_connData, adOpenDynamic, adLockOptimistic)
 
 Call cmbemployé.AddItem(rstEmploye.Fields("Employe"))
 
 cmbemployé.ItemData(cmbemployé.newIndex) = m_iNoEmploye

  Set itmEmploye = lvwEmployes.ListItems.Add

  itmEmploye.Text = rstEmploye.Fields("Employe")
  itmEmploye.Tag = m_iNoEmploye
 
  Call rstEmploye.Close
 
  Call rstAutorisation.Open("SELECT * FROM GrbAutorisationPunch WHERE AutoriserPar = " & m_iNoEmploye, g_connData, adOpenDynamic, adLockOptimistic)
 
  If Not rstAutorisation.EOF Then
  Do While Not rstAutorisation.EOF
  Call rstEmploye.Open("SELECT Employe FROM GrbEmployés WHERE NoEmploye = " & rstAutorisation.Fields("NoEmploye"), g_connData, adOpenDynamic, adLockOptimistic)
 
 Call cmbemployé.AddItem(rstEmploye.Fields("Employe"))
 
cmbemployé.ItemData(cmbemployé.newIndex) = rstAutorisation.Fields("NoEmploye")

 Set itmEmploye = lvwEmployes.ListItems.Add

 itmEmploye.Text = rstEmploye.Fields("Employe")
 itmEmploye.Tag = rstAutorisation.Fields("NoEmploye")
 
 Call rstEmploye.Close
 
 Call rstAutorisation.MoveNext
 Loop

 cmdPunchMultiple.Visible = True
Else
 cmdPunchMultiple.Visible = False 'Gll
End If

1  Call rstAutorisation.Close
Set rstAutorisation = Nothing

 Set rstEmploye = Nothing

If cmbemployé.ListCount = 1 Then
 cmbemployé.ListIndex = 0
End If

 Exit Sub

Oups:

1  wOups "frmPunch", "RemplirComboEmploye", Err, Err.number, Err.Description
End Sub

Private Sub cmdPunchOut_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 On Error GoTo Oups

 If m_bMonthViewHasFocus = True Then
 Call cmdPunchOut_Click
 End If

 Exit Sub

Oups:

 wOups "frmPunch", "cmdPunchOut_MouseUp", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups
 Call Table_exist
 mvwSelection.StartOfWeek = vbSunday
 
 Exit Sub

Oups:

 wOups "frmPunch", "Form_Load", Err, Err.number, Err.Description
End Sub
Private Sub lvwJour_Click()

 On Error GoTo Oups

 If lvwJour.ListItems.count > 0 Then
 If lvwJour.SelectedItem.ForeColor = COLOR_ROUGE Then
 cmdModifierPunchIn.Enabled = True
 cmdModifierPunchOut.Enabled = False
 Else
 cmdModifierPunchIn.Enabled = True
 cmdModifierPunchOut.Enabled = True
 End If
 Else
 cmdModifierPunchIn.Enabled = False
  cmdModifierPunchOut.Enabled = False
  End If

  Exit Sub

Oups:

  wOups "frmPunch", "lvwJour_Click", Err, Err.number, Err.Description
End Sub

Private Sub lvwJour_KeyDown(KeyCode As Integer, Shift As Integer)

 On Error GoTo Oups

 If lvwJour.ListItems.count > 0 Then
 If KeyCode = vbKeyDelete Then
 If MsgBox("Voulez-vous vraiment effacer ce punch ?", vbYesNo) = vbYes Then
 Call EffacerPunch
 End If
 End If
 End If

 Exit Sub

Oups:

 wOups "frmPunch", "lvwJour_KeyDown", Err, Err.number, Err.Description
End Sub

Private Sub EffacerPunch()

 On Error GoTo Oups

 'Efface le punch sélectionné
 Call g_connData.Execute("DELETE * FROM GrbPunch WHERE IDPunch = " & lvwJour.SelectedItem.Tag)
 
 Call RemplirListViewSemaine(False)
 Call RemplirListViewSemaineAutorisation(False)
 Call RemplirListViewJour
 Call RemplirListViewJourAutorisation

 Call CalculerHeureSemaine

 Exit Sub

Oups:

 wOups "frmPunch", "EffacerPunch", Err, Err.number, Err.Description
End Sub

Private Sub mskHeure_GotFocus()

 On Error GoTo Oups

 mskHeure.mask = "##:##"

 Exit Sub

Oups:

 wOups "frmPunch", "mskHeure_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskHeure_LostFocus()

 On Error GoTo Oups

 mskHeure.mask = vbNullString
 
 If mskHeure.Text = "__:__" Then
 mskHeure.Text = vbNullString
 End If

 Exit Sub

Oups:

 wOups "frmPunch", "mskHeure_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskPMHeureDebut_GotFocus()

 On Error GoTo Oups

 mskPMHeureDebut.mask = "##:##"

 Exit Sub

Oups:

 wOups "frmPunch", "mskPMHeureDebut_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskPMHeureDebut_LostFocus()

 On Error GoTo Oups

 mskPMHeureDebut.mask = vbNullString

 If mskPMHeureDebut.Text = "__:__" Then
 mskPMHeureDebut.Text = vbNullString
 End If

 Exit Sub

Oups:

 wOups "frmPunch", "mskPMHeureDebut_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskPMHeureFin_GotFocus()

 On Error GoTo Oups

 mskPMHeureFin.mask = "##:##"

 Exit Sub

Oups:

 wOups "frmPunch", "mskPMHeureFin_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskPMHeureFin_LostFocus()

 On Error GoTo Oups

 mskPMHeureFin.mask = vbNullString

 If mskPMHeureFin.Text = "__:__" Then
 mskPMHeureFin.Text = vbNullString
 End If

 Exit Sub

Oups:

 wOups "frmPunch", "mskPMHeureFin_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskNoProjet_Change()

 On Error GoTo Oups

 If fraPunch.Visible = True Then
 If InStr(1, mskNoProjet.Text, "_") = 0 Then
 Call AfficherTypePunch

 Call AfficherClient
 Else
 txtClient.Text = vbNullString
 End If
 End If

 Call RemplirComboType

 Exit Sub

Oups:

  wOups "frmPunch", "mskNoProjet_Change", Err, Err.number, Err.Description
End Sub

Private Sub mskPMNoProjet_Change()

 On Error GoTo Oups

 If fraPunchMultiple.Visible = True Then
 If InStr(1, mskPMNoProjet.Text, "_") = 0 Then
 Call AfficherTypePunch

 Call AfficherClient
 Else
 txtPMClient.Text = vbNullString
 End If
 End If

 Call RemplirComboType

 Exit Sub

Oups:

  wOups "frmPunch", "mskPMNoProjet_Change", Err, Err.number, Err.Description
End Sub

Private Sub AfficherTypePunch()
 
 On Error GoTo Oups

 Dim sNumero As String
 Dim sType As String
 Dim bPM As Boolean
 
 If fraPunchMultiple.Visible = True Then
 sNumero = mskPMNoProjet.Text
 bPM = True
 Else
 If mskNoProjet.Text <> "_____-__" Then
 sNumero = mskNoProjet.Text
 Else
  sNumero = txtnoprojet.Text
  End If

  bPM = False
  End If
 
64 If Left$(sNumero, 5) = Right$(Year(Date), 1) & "3000" Then
  Select Case Right$(sNumero, 2)
 Case "60": sType = "Petits outils && fourniture"
  Case "70": sType = "Administration de la shop"
  Case "71": sType = "Identification de fils, lamicoïdes, etc."
  Case "72": sType = "Réception de marchandise"
  Case "73": sType = "Support technique informatique et téléphone"
  Case "74": sType = "Commissions"
  Case "75": sType = "Site web && publications"
 Case "76": sType = "Entretien && réparation de la bâtisse"
Case "77": sType = "Ménage général"
 Case "80": sType = "Réparation des outils GRB"
 Case "81": sType = "Lavage des véhicules"
 Case "82": sType = "Entretien && réparation véhicules"
 Case "83": sType = "Formation du personnel"
 Case "85": sType = "Logiciel interne"
 Case "95": sType = "Bâtiment"
 Case "97": sType = "Équipement bureau && informatique"
 Case "99": sType = "Équipements && outillage"
 Case Else: sType = vbNullString
 End Select

If bPM = True Then
 lblPMTypePunch.Caption = sType
 Else
 lblTypePunch.Caption = sType
 End If
Else
 If bPM = True Then
1  lblPMTypePunch.Caption = vbNullString
 Else
 lblTypePunch.Caption = vbNullString
 End If
End If

Exit Sub

Oups:

wOups "frmPunch", "AfficherTypePunch", Err, Err.number, Err.Description
End Sub

Private Sub AfficherClient()

 On Error GoTo Oups

 Dim rstProjSoum As ADODB.Recordset
 Dim rstClient As ADODB.Recordset
 Dim iCompteur As Integer
 Dim sPrefixe As String

 If fraPunchMultiple.Visible = True Then
 If optPMTypePunch(I_OPT_ELECTRIQUE).Value = True Then
 sPrefixe = "E"
 Else
 sPrefixe = "M"
 End If
  Else
  If optTypePunch(I_OPT_ELECTRIQUE).Value = True Then
  sPrefixe = "E"
  Else
  sPrefixe = "M"
  End If
  End If
 
  Set rstProjSoum = New ADODB.Recordset
10 Set rstClient = New ADODB.Recordset
 
If fraPunchMultiple.Visible = True Then
 Call rstProjSoum.Open("SELECT * FROM GrbProjSoum WHERE IDProjSoum = '" & sPrefixe & mskPMNoProjet.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
Else
 Call rstProjSoum.Open("SELECT * FROM GrbProjSoum WHERE IDProjSoum = '" & sPrefixe & mskNoProjet.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
End If
 
If Not rstProjSoum.EOF Then
 Call rstClient.Open("SELECT NomClient FROM GrbClient WHERE IDClient = " & rstProjSoum.Fields("NoClient"), g_connData, adOpenDynamic, adLockOptimistic)
 
 If fraPunchMultiple.Visible = True Then
 txtPMClient.Text = rstClient.Fields("NomClient")
 txtPMClient.Tag = rstProjSoum.Fields("NoClient")
 Else
 txtClient.Text = rstClient.Fields("NomClient")
 txtClient.Tag = rstProjSoum.Fields("NoClient")
 End If
 
 Call rstClient.Close
 Set rstClient = Nothing
 
 If rstProjSoum.Fields("Ouvert") = False Then
 If rstProjSoum.Fields("Type") = "P" Then
1  Call MsgBox("Ce projet est fermé!", vbOKOnly, "Erreur")
 Else
 Call MsgBox("Cette soumission est fermée!", vbOKOnly, "Erreur")
 End If
 End If
Else
 Call MsgBox("Numéro inexistant!", vbOKOnly, "Erreur")

 txtClient.Text = ""
 txtClient.Tag = ""
End If
 
Call rstProjSoum.Close
Set rstProjSoum = Nothing

Exit Sub

Oups:

2  wOups "frmPunch", "AfficherClient", Err, Err.number, Err.Description
End Sub

Private Sub mvwSelection_GotFocus()

'Cette procédure sert à éliminer un bug du controle Active X MonthView
'C'est un bug connu pas Microsoft et la solution suivante est proposée
'Il faut avoir une variable boolean mise à true si le MonthView prend le focus
'et ensuite, en cliquant sur un bouton, si le MonthView a le focus, on force le clique

 On Error GoTo Oups

 m_bMonthViewHasFocus = True

 Exit Sub

Oups:

 wOups "frmPunch", "mvwSelection_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub optHeure_Click(Index As Integer)

 On Error GoTo Oups

 If Index = I_OPT_SYSTEME Then
 mskHeure.Enabled = False
 Else
 mskHeure.Enabled = True

 If m_bModifPunch = False Then
 Call mskHeure.SetFocus
 End If
 End If

 Exit Sub

Oups:

 wOups "frmPunch", "optHeure_Click", Err, Err.number, Err.Description
End Sub

Private Sub lvwJourSemaine_Click(Index As Integer)

 On Error GoTo Oups

 Dim iCompteur As Integer
 Dim sDate As String
 Dim iNbreJour As Integer
 
 'Initialise la couleur en blanc
 For iCompteur = 1 To 7
 lvwJourSemaine(iCompteur).BackColor = &HFFFFFF
 Next
 
 'Sélectionne jour de semaine
 lvwJourSemaine(Index).BackColor = &HE0E0E0

 sDate = lvwJourSemaine(Index).Tag

 Select Case Mid$(sDate, 6, 2)
 Case "01": iNbreJour = 31
 Case "02":
  If CInt(Left$(sDate, 4)) Mod 4 = 0 Then
  iNbreJour = 29
  Else
  iNbreJour = 28
  End If

  Case "03": iNbreJour = 31
  Case "04": iNbreJour = 30
  Case "05": iNbreJour = 31
Case "06": iNbreJour = 30
1 Case "07": iNbreJour = 31
 Case "08": iNbreJour = 31
 Case "09": iNbreJour = 30
 Case "10": iNbreJour = 31
 Case "11": iNbreJour = 30
 Case "12": iNbreJour = 31
End Select

Do While mvwSelection.Day >= iNbreJour
 mvwSelection.Day = mvwSelection.Day - 1
Loop

 'Sélectionne dans calendrier
mvwSelection.Year = Left$(sDate, 4)
1  mvwSelection.Month = Mid$(sDate, 6, 2)
mvwSelection.Day = Right$(sDate, 2)

 'Date choisie
 m_datDateChoisie = DateSerial(mvwSelection.Year, mvwSelection.Month, mvwSelection.Day)

 'Affiche horaire jour
Call RemplirListViewJour
 Call RemplirListViewJourAutorisation

frajour.Visible = True
 fraPunch.Visible = False

1  Call lvwJour.SetFocus

 Exit Sub

Oups:

 woups"frmPunch", "lvwJourSemaine_Click", Err, Erl, "Date cliquée : " & sDate)
End Sub

Private Sub mskNoProjet_KeyPress(KeyAscii As Integer)

 On Error GoTo Oups

 'Pour changer un "m" en un "M"
 If KeyAscii = 10 Then  '10  = m
 KeyAscii = vbKeyM 'M
 End If

 Exit Sub

Oups:

 wOups "frmPunch", "mskNoProjet_KeyPress", Err, Err.number, Err.Description
End Sub

Private Sub mskPMNoProjet_KeyPress(KeyAscii As Integer)

 On Error GoTo Oups

 'Pour changer un "m" en un "M"
 If KeyAscii = 10 Then  '10  = m
 KeyAscii = vbKeyM 'M
 End If

 Exit Sub

Oups:

 wOups "frmPunch", "mskPMNoProjet_KeyPress", Err, Err.number, Err.Description
End Sub

Private Sub cmdOK_Click()

 On Error GoTo Oups

 'Enregistrement du punch in ou du punch out
 Dim rstPunch As ADODB.Recordset
 Dim rstProjSoum As ADODB.Recordset
 Dim sHeure As String
 Dim bModif As Boolean
 Dim iCompteur As Integer
 Dim sPrefixe As String
 Dim sType As String
 Dim sNoProjet As String
 Dim bInstallation As Boolean
 Dim sHeureFin As String
 Dim sNumero As String
 
 
If m_ePunch = I_MODIF_PUNCH_IN Or m_ePunch = I_PUNCH_IN Then
 sNumero = mskNoProjet.Text
 Else

 sNumero = txtnoprojet.Text
 End If
  m_bMonthViewHasFocus = False

  If optHeure(I_OPT_SYSTEME).Value = True Then
  sHeure = GetHeure(Time)
  bModif = False
  Else
  If mskHeure.Text <> vbNullString Then
  If InStr(1, mskHeure.Text, "_") = 0 Then
  sHeure = GetHeure(mskHeure.Text)
 bModif = True
Else
 Call MsgBox("Heure invalide!", vbOKOnly, "Erreur")
 
 Exit Sub
 End If
 Else
 Call MsgBox("Heure invalide!", vbOKOnly, "Erreur")

 Exit Sub
 End If
End If
 
If sHeure <> "" Then
 If m_ePunch = I_MODIF_PUNCH_IN Or m_ePunch = I_PUNCH_IN Then
 If cmbemployé.ListIndex = -1 Or InStr(1, mskNoProjet.Text, "_") > 0 Then
 Call MsgBox("Le nom de l'employé et le numéro de projet sont des champs obligatoires!", vbOKOnly, "Erreur")
 
 Exit Sub
 End If
 End If

 If cmbType.ListIndex = -1 And cmbType.Visible = True Then
 Call MsgBox("Le type est obligatoire!", vbOKOnly, "Erreur")

1  Exit Sub
 End If
 
 'Si c'est une modification de punch in, il faut vérifier l'heure
 'pour être sur qu'elle sont correctes chronologiquement
 If m_ePunch = I_MODIF_PUNCH_IN Then
 If lvwJour.SelectedItem.SubItems(I_LVW_FIN) <> vbNullString Then
 If sHeure > lvwJour.SelectedItem.SubItems(I_LVW_FIN) Then
 Call MsgBox("L'heure de début doit être plus petite que l'heure de fin!", vbOKOnly, "Erreur")
 
 Exit Sub
 End If
 End If
 End If

 'Si c'est une modification de punch in, il faut vérifier l'heure
 'pour être sur qu'elle sont correctes chronologiquement
 If m_ePunch = I_MODIF_PUNCH_OUT Or m_ePunch = I_PUNCH_OUT Then
 If sHeure < lvwJour.SelectedItem.SubItems(I_LVW_DEBUT) Then
 Call MsgBox("L'heure de fin doit être plus grande que l'heure de début!", vbOKOnly, "Erreur")
 
 Exit Sub
 End If
End If

 'Si c'est un punch out avec l'heure de diner et que c'est avant l'heure du diner
 If m_ePunch = I_PUNCH_OUT Or m_ePunch = I_MODIF_PUNCH_OUT Then
 If chkDiner.Value = vbChecked Then
 If optHeureDiner(I_OPT_30_MINUTES).Value = True Then
 If sHeure < "12:30" Then
 Call MsgBox("La case à cocher 'Heure du dîner' ne peut être cochée seulement si l'heure de fin est plus grande que 12:30!", vbOKOnly, "Erreur")

 Exit Sub
 Else
 If lvwJour.SelectedItem.SubItems(I_LVW_DEBUT) > "12:00" Then
 Call MsgBox("La case à cocher 'Heure du dîner' ne peut être cochée seulement si l'heure de début est plus petite que 12:00!", vbOKOnly, "Erreur")

 Exit Sub
 End If
 End If
 Else
 If sHeure < "13:00" Then
 Call MsgBox("La case à cocher 'Heure du dîner' ne peut être cochée seulement si l'heure de fin est plus grande que 13:00!", vbOKOnly, "Erreur")

 Exit Sub
 Else
 If lvwJour.SelectedItem.SubItems(I_LVW_DEBUT) > "12:00" Then
 Call MsgBox("La case à cocher 'Heure du dîner' ne peut être cochée seulement si l'heure de début est plus petite que 12:00!", vbOKOnly, "Erreur")

 Exit Sub
 End If
 End If
 End If
 End If
 End If
 
If optTypePunch(I_OPT_ELECTRIQUE).Value = True Then
4 sPrefixe = "E"
4 Else
4 sPrefixe = "M"
4 End If
 
4 If m_ePunch = I_MODIF_PUNCH_IN Or m_ePunch = I_PUNCH_IN Then
4 sNoProjet = mskNoProjet.Text
4 Else
4 sNoProjet = txtnoprojet.Text
4 End If

4 If cmbType.Visible = True Then
4 If IsNumeric(Right$(sNoProjet, 2)) Then
4  If CInt(Right$(sNoProjet, 2)) >= 51 And CInt(Right$(sNoProjet, 2)) <= 5 Then
4  bInstallation = True
4  Else
4  bInstallation = False
4  End If
4  Else
4  bInstallation = False
4  End If
 
50 If bInstallation = True Then
 If sPrefixe = "E" Then
 Select Case cmbType.ListIndex
 Case I_TYPE_ELEC_INSTALLATION: sType = "Installation"
 Case I_TYPE_ELEC_MISE_SERVICE: sType = "MiseService"
 End Select
 Else
 Select Case cmbType.ListIndex
 Case I_TYPE_MEC_INSTALLATION: sType = "Installation"
 End Select
 End If
 Else
 If sPrefixe = "E" Then
 sType = cmbType.Text
610
  Else
620
 sType = cmbType.Text
670
6  End If
6  End If
6  End If

6  If m_ePunch = I_MODIF_PUNCH_IN Or m_ePunch = I_PUNCH_IN Then
6  Set rstProjSoum = New ADODB.Recordset

70 Call rstProjSoum.Open("SELECT * FROM GrbProjSoum WHERE IDProjSoum = '" & sPrefixe & mskNoProjet.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
  If Not rstProjSoum.EOF Then
  If txtClient.Text <> "" Then
  If rstProjSoum.Fields("Ouvert") = False Then
  If rstProjSoum.Fields("Type") = "P" Then
  Call MsgBox("Ce projet est fermé!", vbOKOnly, "Erreur")
  Else
  Call MsgBox("Cette soumission est fermée!", vbOKOnly, "Erreur")
  End If
 
  Call rstProjSoum.Close
  Set rstProjSoum = Nothing
 
  Exit Sub
   End If
   Else
7  Call MsgBox("Le client ne doit pas être vide!", vbOKOnly, "Erreur")

7  Call rstProjSoum.Close
7  Set rstProjSoum = Nothing
 
7  Exit Sub
7  End If
7  Else
80 Call MsgBox("Numéro inexistant!", vbOKOnly, "Erreur")
 
  Call rstProjSoum.Close
  Set rstProjSoum = Nothing
 
  Exit Sub
  End If
 
  Call rstProjSoum.Close
  Set rstProjSoum = Nothing
  End If

  If Trim$(txtCommentaires.Text) = "" Then
  Call MsgBox("Le commentaire est obligatoire!", vbOKOnly, "Erreur")

  Exit Sub
  End If
 
   Set rstPunch = New ADODB.Recordset
 
 'Selon le mode
   Select Case m_ePunch
 'Si c'est un punch in
 Case I_PUNCH_IN:
 'On ouvre le recordset avec la date et le no d'employé
   Call rstPunch.Open("SELECT * FROM GrbPunch WHERE NoEmploye = " & cmbemployé.ItemData(cmbemployé.ListIndex) & " AND Date = '" & ConvertDate(m_datDateChoisie) & "' ORDER BY Date,HeureDébut", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Si il y a des enregistrements
   If Not rstPunch.EOF Then
 'On va au dernier
8  Call rstPunch.MoveLast
 
 'On vérifie si le dernier punch out n'a pas été fait
8  If IsNull(rstPunch.Fields("HeureFin")) Or rstPunch.Fields("HeureFin") = vbNullString Then
 'On fait le punch out
8  rstPunch.Fields("ModifFin") = bModif
8  rstPunch.Fields("HeureFin") = sHeure

90 Call rstPunch.Update
  End If
  End If
 
  Call rstPunch.AddNew
 
  rstPunch.Fields("NoEmploye") = cmbemployé.ItemData(cmbemployé.ListIndex)
  rstPunch.Fields("NoProjet") = sPrefixe & mskNoProjet.Text
  rstPunch.Fields("Date") = ConvertDate(DateSerial(mvwSelection.Year, mvwSelection.Month, mvwSelection.Day))
  rstPunch.Fields("ModifDébut") = bModif
  rstPunch.Fields("HeureDébut") = sHeure
  rstPunch.Fields("Commentaire") = txtCommentaires.Text
  rstPunch.Fields("NoClient") = txtClient.Tag

  If chkKM.Value = vbChecked Then
 rstPunch.Fields("KM") = True

   If txtKM.Text <> "" Then
 txtKM.Text = Replace(txtKM.Text, ".", ",")

   If IsNumeric(txtKM.Text) Then
 rstPunch.Fields("NbreKM") = txtKM.Text
   Else
 rstPunch.Fields("NbreKM") = 0
9  End If
 Else
 rstPunch.Fields("KM") = False
 rstPunch.Fields("NbreKM") = ""
 End If
 Else
 rstPunch.Fields("KM") = False
 rstPunch.Fields("NbreKM") = ""
 End If
 If Mid$(sNumero, 2, 1) = "1" Then
 rstPunch.Fields("Type") = "Soumission"
 Else
 rstPunch.Fields("Type") = sType
 End If
 Call rstPunch.Update
 
 Call rstPunch.Close
 
 'Si c'est un punch out
1 Case I_PUNCH_OUT:
 'Si l'élément choisi est en noir, le punch out a déjà été fait
10  If lvwJour.SelectedItem.ForeColor = COLOR_ROUGE Then
 'On ouvre le recordset avec le numéro de punch
10  Call rstPunch.Open("SELECT * FROM GrbPunch WHERE IDPunch = " & lvwJour.SelectedItem.Tag, g_connData, adOpenDynamic, adLockOptimistic)
 
10  If chkDiner.Value = vbChecked Then
10  rstPunch.Fields("ModifFin") = False
10  rstPunch.Fields("HeureFin") = "12:00"
10  Else
10  rstPunch.Fields("ModifFin") = bModif
10  rstPunch.Fields("HeureFin") = sHeure
1 End If

1 rstPunch.Fields("Commentaire") = txtCommentaires.Text
 
1 If chkKM.Value = vbChecked Then
1 rstPunch.Fields("KM") = True

1 If txtKM.Text <> "" Then
1 txtKM.Text = Replace(txtKM.Text, ".", ",")

1 If IsNumeric(txtKM.Text) Then
1 rstPunch.Fields("NbreKM") = txtKM.Text
1 Else
1 rstPunch.Fields("NbreKM") = 0
1 End If
1 Else
1 rstPunch.Fields("KM") = False
1 rstPunch.Fields("NbreKM") = ""
 End If
1 Else
 rstPunch.Fields("KM") = False
1 rstPunch.Fields("NbreKM") = ""
 End If
 If Mid$(sNumero, 2, 1) = "1" Then
 rstPunch.Fields("Type") = "Soumission"
 Else
11  rstPunch.Fields("Type") = sType
 End If
 Call rstPunch.Update

 If chkDiner.Value = vbChecked Then
1 Call rstPunch.AddNew
 
1 rstPunch.Fields("NoEmploye") = cmbemployé.ItemData(cmbemployé.ListIndex)

1 If mskNoProjet.Text = "_____-__" Then
1 rstPunch.Fields("NoProjet") = sPrefixe & txtnoprojet.Text
1 Else
1 rstPunch.Fields("NoProjet") = sPrefixe & mskNoProjet.Text
1 End If
 
1 rstPunch.Fields("Date") = ConvertDate(DateSerial(mvwSelection.Year, mvwSelection.Month, mvwSelection.Day))
1 rstPunch.Fields("ModifDébut") = False

1 If optHeureDiner(I_OPT_30_MINUTES).Value = True Then
1 rstPunch.Fields("HeureDébut") = "12:30"
1 Else
1 rstPunch.Fields("HeureDébut") = "13:00"
1 End If

1 rstPunch.Fields("Commentaire") = txtCommentaires.Text
1 rstPunch.Fields("NoClient") = txtClient.Tag
1 rstPunch.Fields("ModifFin") = bModif
1 rstPunch.Fields("HeureFin") = sHeure
 If Mid$(sNumero, 2, 1) = "1" Then
 rstPunch.Fields("Type") = "Soumission"
 Else
1 rstPunch.Fields("Type") = sType
 End If
1 Call rstPunch.Update
1 End If
 
1 Call rstPunch.Close
1 Else
1 Call MsgBox("Le punch out a déjà été fait!", vbOKOnly, "Erreur")
1 End If
 
 'Si c'est une modification de punch in
1 Case I_MODIF_PUNCH_IN:
 'On ouvre le recordset avec le numéro de punch
1 Call rstPunch.Open("SELECT * FROM GrbPunch WHERE IDPunch = " & lvwJour.SelectedItem.Tag, g_connData, adOpenDynamic, adLockOptimistic)
 
1 rstPunch.Fields("NoEmploye") = cmbemployé.ItemData(cmbemployé.ListIndex)
1 rstPunch.Fields("NoProjet") = sPrefixe & mskNoProjet.Text
1 rstPunch.Fields("NoClient") = txtClient.Tag
1 rstPunch.Fields("Date") = ConvertDate(DateSerial(mvwSelection.Year, mvwSelection.Month, mvwSelection.Day))

1 If bModif = True Then
1 If rstPunch.Fields("HeureDébut") <> sHeure Then
1 rstPunch.Fields("ModifDébut") = True
1 Else
1 rstPunch.Fields("ModifDébut") = False
1 End If
1 Else
1 rstPunch.Fields("ModifDébut") = False
14 End If
 
14 rstPunch.Fields("HeureDébut") = sHeure

14 rstPunch.Fields("Commentaire") = txtCommentaires.Text
 
14 If chkKM.Value = vbChecked Then
14 rstPunch.Fields("KM") = True

14 If txtKM.Text <> "" Then
14 txtKM.Text = Replace(txtKM.Text, ".", ",")

14 If IsNumeric(txtKM.Text) Then
14 rstPunch.Fields("NbreKM") = txtKM.Text
14 Else
14 rstPunch.Fields("NbreKM") = 0
14  End If
14  Else
14  rstPunch.Fields("KM") = False
14  rstPunch.Fields("NbreKM") = 0
14  End If
14  Else
14  rstPunch.Fields("KM") = False
14  rstPunch.Fields("NbreKM") = ""
150 End If
 If Mid$(sNumero, 2, 1) = "1" Then
 rstPunch.Fields("Type") = "Soumission"
 Else
1 rstPunch.Fields("Type") = sType
 End If
 Call rstPunch.Update
 
 Call rstPunch.Close
 
 'Si c'est une modification de punch out
 Case I_MODIF_PUNCH_OUT:
 'On ouvre le recordset avec le numéro de punch
 Call rstPunch.Open("SELECT * FROM GrbPunch WHERE IDPunch = " & lvwJour.SelectedItem.Tag, g_connData, adOpenDynamic, adLockOptimistic)
 
 If chkDiner.Value = vbChecked Then
 sHeureFin = rstPunch.Fields("HeureFin")

 rstPunch.Fields("ModifFin") = False
 rstPunch.Fields("HeureFin") = "12:00"
 Else
 If bModif = True Then
15  If rstPunch.Fields("HeureFin") <> sHeure Then
15  rstPunch.Fields("ModifFin") = True
15  Else
15  rstPunch.Fields("ModifFin") = False
15  End If
15  Else
15  rstPunch.Fields("ModifFin") = False
15  End If

160 rstPunch.Fields("HeureFin") = sHeure
 End If

 rstPunch.Fields("Commentaire") = txtCommentaires.Text
 
 If chkKM.Value = vbChecked Then
 rstPunch.Fields("KM") = True

 If txtKM.Text <> "" Then
 txtKM.Text = Replace(txtKM.Text, ".", ",")

 If IsNumeric(txtKM.Text) Then
 rstPunch.Fields("NbreKM") = txtKM.Text
 Else
 rstPunch.Fields("NbreKM") = 0
 End If
16  Else
16  rstPunch.Fields("KM") = False
16  rstPunch.Fields("NbreKM") = ""
16  End If
16  Else
16  rstPunch.Fields("KM") = False
16  rstPunch.Fields("NbreKM") = ""
16  End If

 If Mid$(sNumero, 2, 1) = "1" Then
 rstPunch.Fields("Type") = "Soumission"
 Else
170 rstPunch.Fields("Type") = sType
 End If
 Call rstPunch.Update

 If chkDiner.Value = vbChecked Then
 Call rstPunch.AddNew
 
 rstPunch.Fields("NoEmploye") = cmbemployé.ItemData(cmbemployé.ListIndex)

 rstPunch.Fields("NoProjet") = sPrefixe & txtnoprojet.Text

 rstPunch.Fields("Date") = ConvertDate(DateSerial(mvwSelection.Year, mvwSelection.Month, mvwSelection.Day))
 rstPunch.Fields("ModifDébut") = False

 If optHeureDiner(I_OPT_30_MINUTES).Value = True Then
 rstPunch.Fields("HeureDébut") = "12:30"
 Else
 rstPunch.Fields("HeureDébut") = "13:00"
1   End If

1   rstPunch.Fields("Commentaire") = txtCommentaires.Text
17  rstPunch.Fields("NoClient") = txtClient.Tag

17  If bModif = True Then
17  If rstPunch.Fields("HeureFin") <> sHeureFin Then
17  rstPunch.Fields("ModifFin") = True
17  Else
17  rstPunch.Fields("ModifFin") = False
180 End If
 Else
 rstPunch.Fields("ModifFin") = False
 End If

 rstPunch.Fields("HeureFin") = sHeure

 rstPunch.Fields("Type") = sType

 Call rstPunch.Update
 End If
 
 Call rstPunch.Close
 End Select
 
 Set rstPunch = Nothing
 
 Call RemplirListViewSemaine(True)
1   Call RemplirListViewSemaineAutorisation(True)
1   Call RemplirListViewJour
1   Call RemplirListViewJourAutorisation

1   Call CalculerHeureSemaine

18  frajour.Visible = True
18  fraPunch.Visible = False
18  fraPunchMultiple.Visible = False
18  End If

190Exit Sub

Oups:

wOups "frmPunch", "cmdOK_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdPMOK_Click()

 On Error GoTo Oups

 'Enregistrement du punch in ou du punch out
 Dim rstPunch As ADODB.Recordset
 Dim rstProjSoum As ADODB.Recordset
 Dim sHeureDebut As String
 Dim sHeureFin As String
 Dim iCompteur As Integer
 Dim bSelect As Boolean
 Dim sPrefixe As String
 Dim sType As String
 Dim bInstallation As Boolean

 m_bMonthViewHasFocus = False

  If mskPMHeureDebut.Text <> vbNullString Then
  If InStr(1, mskPMHeureDebut.Text, "_") = 0 Then
  sHeureDebut = GetHeure(mskPMHeureDebut.Text)
  Else
  Call MsgBox("Heure de début invalide!", vbOKOnly, "Erreur")
 
  Exit Sub
  End If
  Else
Call MsgBox("Heure de début invalide!", vbOKOnly, "Erreur")

1 Exit Sub
End If
 
If mskPMHeureFin.Text <> vbNullString Then
 If InStr(1, mskPMHeureFin.Text, "_") = 0 Then
 sHeureFin = GetHeure(mskPMHeureFin.Text)
 Else
 Call MsgBox("Heure de fin invalide!", vbOKOnly, "Erreur")
 
 Exit Sub
 End If
Else
 Call MsgBox("Heure de fin invalide!", vbOKOnly, "Erreur")

Exit Sub
End If

 If cmbPMType.ListIndex = -1 And cmbPMType.Visible = True Then
 Call MsgBox("Le type est obligatoire!", vbOKOnly, "Erreur")

 Exit Sub
End If

 For iCompteur = 1 To lvwEmployes.ListItems.count
1  If lvwEmployes.ListItems(iCompteur).Checked = True Then
 bSelect = True

 Exit For
 End If
Next

If bSelect = False Then
 Call MsgBox("Vous devez choisir au moins 1 employé!", vbOKOnly, "Erreur")

 Exit Sub
End If

If InStr(1, mskPMNoProjet.Text, "_") > 0 Then
 Call MsgBox("Le numéro de projet est obligatoire!", vbOKOnly, "Erreur")
 
 Exit Sub
End If

2  If sHeureDebut > sHeureFin Then
 Call MsgBox("L'heure de début doit être plus petite que l'heure de fin!", vbOKOnly, "Erreur")
 
Exit Sub
End If

 'Si c'est un punch out avec l'heure de diner et que c'est avant l'heure du diner
2  If chkPMDiner.Value = vbChecked Then
 If optPMHeureDiner(I_OPT_30_MINUTES).Value = True Then
 If sHeureDebut > "12:00" Or sHeureFin < "12:30" Then
 Call MsgBox("La case à cocher 'Heure du dîner' ne peut être cochée que si l'heure de début est plus petite que 12:00" & vbNewLine & _
 " ou que l'heure de fin est plus grande que 12:30!", vbOKOnly, "Erreur")

 Exit Sub
End If
 Else
 If sHeureDebut > "12:00" Or sHeureFin < "13:00" Then
 Call MsgBox("La case à cocher 'Heure du dîner' ne peut être cochée que si l'heure de début est plus petite que 12:00" & vbNewLine & _
 " ou que l'heure de fin est plus grande que 13:00!", vbOKOnly, "Erreur")

 Exit Sub
 End If
 End If
End If

If optPMTypePunch(I_OPT_ELECTRIQUE).Value = True Then
 sPrefixe = "E"
Else
sPrefixe = "M"
End If
 
3  If cmbPMType.Visible = True Then
 If IsNumeric(Right$(mskPMNoProjet.Text, 2)) Then
 If CInt(Right$(mskPMNoProjet.Text, 2)) >= 51 And CInt(Right$(mskPMNoProjet.Text, 2)) <= 5 Then
 bInstallation = True
 Else
 bInstallation = False
 End If
4 Else
4 bInstallation = False
4 End If
 
4 If bInstallation = True Then
4 If sPrefixe = "E" Then
4 Select Case cmbPMType.ListIndex
 Case I_TYPE_ELEC_INSTALLATION: sType = "Installation"
4 Case I_TYPE_ELEC_MISE_SERVICE: sType = "Formation"
4 End Select
4 Else
4 Select Case cmbPMType.ListIndex
 Case I_TYPE_MEC_INSTALLATION: sType = "Installation"
4 End Select
4  End If
4  Else
4  If sPrefixe = "E" Then
 sType = cmbPMType.Text
480
530
 Else
 sType = cmbPMType.Text

590
5  End If
60 End If
60 End If

  Set rstProjSoum = New ADODB.Recordset

  Call rstProjSoum.Open("SELECT * FROM GrbProjSoum WHERE IDProjSoum = '" & sPrefixe & mskPMNoProjet.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
  If Not rstProjSoum.EOF Then
  If txtPMClient.Text <> "" Then
  If rstProjSoum.Fields("Ouvert") = False Then
  If rstProjSoum.Fields("Type") = "P" Then
  Call MsgBox("Ce projet est fermé!", vbOKOnly, "Erreur")
  Else
  Call MsgBox("Cette soumission est fermée!", vbOKOnly, "Erreur")
  End If
 
6  Call rstProjSoum.Close
6  Set rstProjSoum = Nothing
 
6  Exit Sub
6  End If
6  Else
6  Call MsgBox("Le client ne doit pas être vide!", vbOKOnly, "Erreur")

6  Call rstProjSoum.Close
6  Set rstProjSoum = Nothing

70 Exit Sub
  End If
  Else
  Call MsgBox("Numéro inexistant!", vbOKOnly, "Erreur")

  Call rstProjSoum.Close
  Set rstProjSoum = Nothing
 
  Exit Sub
  End If
 
  Call rstProjSoum.Close
  Set rstProjSoum = Nothing

  If Trim$(txtPMCommentaire.Text) = "" Then
  Call MsgBox("Le commentaire est obligatoire!", vbOKOnly, "Erreur")

   Exit Sub
   End If
 
7  Set rstPunch = New ADODB.Recordset
 
7  For iCompteur = 1 To lvwEmployes.ListItems.count
7  If lvwEmployes.ListItems(iCompteur).Checked = True Then
7  Call rstPunch.Open("SELECT * FROM GrbPunch", g_connData, adOpenDynamic, adLockOptimistic)
 
7  Call rstPunch.AddNew
 
7  rstPunch.Fields("NoEmploye") = lvwEmployes.ListItems(iCompteur).Tag
80 rstPunch.Fields("NoProjet") = sPrefixe & mskPMNoProjet.Text
  rstPunch.Fields("Date") = ConvertDate(DateSerial(mvwSelection.Year, mvwSelection.Month, mvwSelection.Day))
  rstPunch.Fields("ModifDébut") = False
  rstPunch.Fields("HeureDébut") = sHeureDebut

  rstPunch.Fields("ModifFin") = False

  If chkPMDiner.Value = vbChecked Then
  rstPunch.Fields("HeureFin") = "12:00"
  Else
  rstPunch.Fields("HeureFin") = sHeureFin
  End If

  rstPunch.Fields("Commentaire") = txtPMCommentaire.Text
  rstPunch.Fields("NoClient") = txtPMClient.Tag

   rstPunch.Fields("KM") = False
   rstPunch.Fields("NbreKM") = ""

   rstPunch.Fields("Type") = sType
 
   Call rstPunch.Update

8  If chkPMDiner.Value = vbChecked Then
8  Call rstPunch.AddNew
 
8  rstPunch.Fields("NoEmploye") = lvwEmployes.ListItems(iCompteur).Tag
8  rstPunch.Fields("NoProjet") = sPrefixe & mskPMNoProjet.Text
90 rstPunch.Fields("Date") = ConvertDate(DateSerial(mvwSelection.Year, mvwSelection.Month, mvwSelection.Day))
  rstPunch.Fields("ModifDébut") = False

  If optPMHeureDiner(I_OPT_30_MINUTES).Value = True Then
  rstPunch.Fields("HeureDébut") = "12:30"
  Else
  rstPunch.Fields("HeureDébut") = "13:00"
  End If

  rstPunch.Fields("Commentaire") = txtPMCommentaire.Text
  rstPunch.Fields("NoClient") = txtPMClient.Tag
  rstPunch.Fields("ModifFin") = False
  rstPunch.Fields("HeureFin") = sHeureFin

  rstPunch.Fields("Type") = sType

 Call rstPunch.Update
   End If
 
 Call rstPunch.Update
 
   Call rstPunch.Close
 End If
   Next

 Set rstPunch = Nothing

9  Call RemplirListViewSemaine(True)
 Call RemplirListViewSemaineAutorisation(True)
100 Call RemplirListViewJour
10 Call RemplirListViewJourAutorisation

10 Call CalculerHeureSemaine

frajour.Visible = True
10 fraPunch.Visible = False
fraPunchMultiple.Visible = False

10 Exit Sub

Oups:

wOups "frmPunch", "cmdPMOK_Click", Err, Err.number, Err.Description
End Sub

Private Function GetHeure(ByVal sHeure As String) As String

 On Error GoTo Oups

 Dim datHeure As Date
 Dim b24Heure As Boolean

 If IsNumeric(Left$(sHeure, 2)) And IsNumeric(Mid$(sHeure, 4, 2)) Then
 If (CInt(Left$(sHeure, 2)) < 0 Or CInt(Left$(sHeure, 2)) > 24) Or (CInt(Mid$(sHeure, 4, 2)) < 0 Or CInt(Mid$(sHeure, 4, 2)) > 59) Then
 Call MsgBox("Heure incorrecte!", vbOKOnly, "Erreur")

 Exit Function
 End If
 Else
 Call MsgBox("Heure incorrecte!", vbOKOnly, "Erreur")

 Exit Function
  End If

  sHeure = Replace(sHeure, "AM", "")

  If InStr(1, sHeure, "PM") > 0 Then
  sHeure = Trim$(Replace(sHeure, "PM", ""))

  sHeure = CInt(Left$(sHeure, 2)) + 12 & Right$(sHeure, Len(sHeure) - 2)

  b24Heure = True
  End If

  sHeure = Left$(sHeure, 5)

10 If sHeure <> "24:00" Then
1 datHeure = CDate(sHeure)

 If Minute(datHeure) <= 5 Then
 datHeure = TimeSerial(Hour(datHeure), 0, 0)
 Else
 If Minute(datHeure) <= 24 Then
 datHeure = TimeSerial(Hour(datHeure), 15, 0)
 Else
 If Minute(datHeure) <= 35 Then
 datHeure = TimeSerial(Hour(datHeure), 30, 0)
 Else
 If Minute(datHeure) <= 54 Then
 datHeure = TimeSerial(Hour(datHeure), 45, 0)
 Else
 datHeure = TimeSerial(Hour(datHeure) + 1, 0, 0)
 End If
 End If
 End If
 End If

1  GetHeure = Right$("0" & Hour(datHeure), 2) & ":" & Right$("0" & Minute(datHeure), 2)
 Else
 GetHeure = sHeure
End If

Exit Function

Oups:

wOups "frmPunch", "GetHeure", Err, Err.number, Err.Description
End Function

Private Sub optHeure_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

 On Error GoTo Oups

 If m_bMonthViewHasFocus = True Then
 If optHeure(Index).Value = False Then
 optHeure(Index).Value = True
 End If
 End If

 Exit Sub

Oups:

 wOups "frmPunch", "optHeure_MouseUp", Err, Err.number, Err.Description
End Sub

Private Sub optTypePunch_Click(Index As Integer)
 
 On Error GoTo Oups

 If InStr(1, mskNoProjet.Text, "_") = 0 Then
 Call AfficherTypePunch

 Call AfficherClient
 Else
 If fraPunch.Visible = True Then
 Call mskNoProjet.SetFocus
 End If
 End If

 Select Case Index
 Case I_OPT_ELECTRIQUE: lblPrefixe.Caption = "E"
 Case I_OPT_MECANIQUE: lblPrefixe.Caption = "M"
  End Select

  Call RemplirComboType

  m_bMonthViewHasFocus = False

  Exit Sub

Oups:

  wOups "frmPunch", "optTypePunch_Click", Err, Err.number, Err.Description
End Sub

Private Sub optTypePunch_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

 On Error GoTo Oups

 If m_bMonthViewHasFocus = True Then
 Call optTypePunch_Click(Index)
 End If

 Exit Sub

Oups:

 wOups "frmPunch", "optTypePunch_MouseUp", Err, Err.number, Err.Description
End Sub

Private Sub optPMTypePunch_Click(Index As Integer)
 
 On Error GoTo Oups

 If InStr(1, mskPMNoProjet.Text, "_") = 0 Then
 Call AfficherTypePunch

 Call AfficherClient
 Else
 If fraPunchMultiple.Visible = True Then
 Call mskPMNoProjet.SetFocus
 End If
 End If

 Select Case Index
 Case I_OPT_ELECTRIQUE: lblPMPrefixe.Caption = "E"
 Case I_OPT_MECANIQUE: lblPMPrefixe.Caption = "M"
  End Select

  Call RemplirComboType

  m_bMonthViewHasFocus = False

  Exit Sub

Oups:

  wOups "frmPunch", "optPMTypePunch_Click", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboType()
 
 On Error GoTo Oups

 Dim cmbSource As ComboBox
 Dim lblSource As Label
 Dim sType As String
 Dim sNumero As String
 Dim bInstallation As Boolean
 Dim tblremplircombo As ADODB.Recordset
 Set tblremplircombo = New ADODB.Recordset
 
 If fraPunchMultiple.Visible = True Then
 
 Set cmbSource = cmbPMType
 Set lblSource = lblPMType
 
 sType = lblPMPrefixe.Caption

 sNumero = mskPMNoProjet.Text
  Else
 
  Set cmbSource = cmbType
  Set lblSource = lblType
 
  sType = lblPrefixe.Caption

  If m_ePunch = I_MODIF_PUNCH_IN Or m_ePunch = I_PUNCH_IN Then

  sNumero = mskNoProjet.Text
  Else

  sNumero = txtnoprojet.Text
End If
End If
 
Call cmbSource.Clear

If Mid$(sNumero, 2, 1) = "1" Or Mid$(sNumero, 2, 4) = "3000" Then
 cmbSource.Visible = False
 lblSource.Visible = False

 Exit Sub
Else

 cmbSource.Visible = True
 lblSource.Visible = True
End If

If IsNumeric(Right$(sNumero, 2)) Then

If CInt(Right$(sNumero, 2)) >= 51 And CInt(Right$(sNumero, 2)) <= 5 Then

 bInstallation = True
 Else

 bInstallation = False
 End If
Else

 bInstallation = False
1  End If
 
 If bInstallation = True Then

 If sType = "E" Then

 Call cmbSource.AddItem("Installation")
 Call cmbSource.AddItem("Mise en service")
 Else

 Call cmbSource.AddItem("Installation")
 End If
Else
 If sType = "E" Then

 Call tblremplircombo.Open("Select * from TBL_Punch_Type WHERE MODE = 'E' Order by name ", g_connData, adOpenDynamic, adLockOptimistic)
 Do While Not tblremplircombo.EOF
 cmbSource.AddItem (tblremplircombo.Fields("name"))
 Call tblremplircombo.MoveNext
275
 Loop
 Call tblremplircombo.Close
 Set tblremplircombo = Nothing
295
 
Else

3 Call tblremplircombo.Open("Select * from TBL_Punch_Type WHERE MODE = 'M' Order by name ", g_connData, adOpenDynamic, adLockOptimistic)
 Do While Not tblremplircombo.EOF
 cmbSource.AddItem (tblremplircombo.Fields("name"))
 Call tblremplircombo.MoveNext
 Loop
 Call tblremplircombo.Close
 Set tblremplircombo = Nothing
340
345
350
 
 End If
3  End If
 
Set cmbSource = Nothing

3  Exit Sub

Oups:

wOups "frmPunch", "RemplirComboType", Err, Err.number, Err.Description
End Sub

Private Sub optPMTypePunch_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

 On Error GoTo Oups

 If m_bMonthViewHasFocus = True Then
 Call optPMTypePunch_Click(Index)
 End If

 Exit Sub

Oups:

 wOups "frmPunch", "optPMTypePunch_MouseUp", Err, Err.number, Err.Description
End Sub

Private Sub txtKM_LostFocus()
 
 On Error GoTo Oups

 txtKM.Text = Replace(txtKM.Text, ".", ",")

 Exit Sub
 
Oups:

 wOups "frmPunch", "txtKM_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mvwSelection_SelChange(ByVal StartDate As Date, ByVal EndDate As Date, Cancel As Boolean)

 On Error GoTo Oups

 If Month(m_datDateChoisie) <> mvwSelection.Month Or _
 Year(m_datDateChoisie) <> mvwSelection.Year Or _
 Day(m_datDateChoisie) <> mvwSelection.Day Then
 Call AfficherDate

 Call CalculerHeureSemaine
 End If

 Exit Sub

Oups:

 wOups "frmPunch", "mvwSelection_SelChange", Err, Err.number, Err.Description
End Sub

Private Function VerifierModificationDate() As Boolean
 
 On Error GoTo Oups

 Dim bModif As Boolean
 Dim datSelected As Date
 Dim datToday As Date
 Dim datFirstDaySelected As Date
 Dim datFirstDayToday As Date

 datSelected = mvwSelection.Value
 datToday = Date
 datFirstDaySelected = GetFirstDay(datSelected)
 datFirstDayToday = GetFirstDay(datToday)

 If g_bPunchSemaineAnterieure = False Then
  If datSelected <> datToday Then
  If Weekday(datToday, vbSunday) = vbSunday Or _
 Weekday(datToday, vbSunday) = vbMonday Then
  If (datFirstDaySelected = datFirstDayToday) Or DateDiff("d", datFirstDaySelected, datFirstDayToday) =   Then
  bModif = True
  Else
  bModif = False
  End If
  Else
 If datFirstDaySelected = datFirstDayToday Then
 bModif = True
 End If
 End If
 Else
 bModif = True
 End If
Else
 bModif = True
End If

If bModif = False Then
 Call MsgBox("Impossible de modifier les punchs de cette journée!", vbOKOnly, "Erreur")
1  End If

VerifierModificationDate = bModif

 Exit Function

Oups:

wOups "frmPunch", "VerifierModificationDate", Err, Err.number, Err.Description
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


