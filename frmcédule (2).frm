VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCédule 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cédule"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10950
   ClipControls    =   0   'False
   Icon            =   "frmcédule.frx":0000
   LinkTopic       =   "frmcédule"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   10950
   Begin MSComCtl2.MonthView mvwSelection 
      Height          =   4020
      Left            =   240
      TabIndex        =   39
      Top             =   840
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
   Begin VB.ComboBox cmbfinprojet 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3120
      TabIndex        =   1
      Text            =   "cmbfinprojet"
      Top             =   360
      Width           =   1815
   End
   Begin VB.Frame frasemaine 
      BackColor       =   &H00404040&
      Height          =   2655
      Left            =   0
      TabIndex        =   40
      Top             =   4920
      Width           =   10935
      Begin MSComctlLib.ListView lstjoursemaine 
         Height          =   1900
         Index           =   1
         Left            =   0
         TabIndex        =   55
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "no"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "nom"
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "heure"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView lstjoursemaine 
         Height          =   1900
         Index           =   2
         Left            =   1560
         TabIndex        =   56
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "no"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "nom"
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "heure"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView lstjoursemaine 
         Height          =   1900
         Index           =   3
         Left            =   3120
         TabIndex        =   57
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "no"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "nom"
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "heure"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView lstjoursemaine 
         Height          =   1900
         Index           =   4
         Left            =   4680
         TabIndex        =   58
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "no"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "nom"
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "heure"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView lstjoursemaine 
         Height          =   1900
         Index           =   5
         Left            =   6240
         TabIndex        =   59
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "no"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "nom"
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "heure"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView lstjoursemaine 
         Height          =   1900
         Index           =   6
         Left            =   7800
         TabIndex        =   60
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "no"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "nom"
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "heure"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView lstjoursemaine 
         Height          =   1900
         Index           =   7
         Left            =   9360
         TabIndex        =   61
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "no"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "nom"
            Object.Width           =   794
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "heure"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label lbljour 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   10030
         TabIndex        =   53
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lbljour 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   8470
         TabIndex        =   51
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lbljour 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   6910
         TabIndex        =   49
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lbljour 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   5350
         TabIndex        =   47
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lbljour 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   3790
         TabIndex        =   46
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lbljour 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   2230
         TabIndex        =   43
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lbljour 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   670
         TabIndex        =   42
         Top             =   360
         Width           =   495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   1560
         X2              =   1560
         Y1              =   360
         Y2              =   2520
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         X1              =   9360
         X2              =   9360
         Y1              =   360
         Y2              =   2520
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         X1              =   7800
         X2              =   7800
         Y1              =   360
         Y2              =   2520
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   6240
         X2              =   6240
         Y1              =   360
         Y2              =   2520
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   4680
         X2              =   4680
         Y1              =   360
         Y2              =   2520
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   3120
         X2              =   3120
         Y1              =   360
         Y2              =   2520
      End
      Begin VB.Label lbljourstr 
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
         Left            =   9480
         TabIndex        =   54
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lbljourstr 
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
         Left            =   7920
         TabIndex        =   52
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lbljourstr 
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
         Left            =   6360
         TabIndex        =   50
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lbljourstr 
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
         Left            =   4800
         TabIndex        =   48
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lbljourstr 
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
         Left            =   3240
         TabIndex        =   45
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lbljourstr 
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
         Left            =   1680
         TabIndex        =   44
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lbljourstr 
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
         TabIndex        =   41
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame fraliste 
      BackColor       =   &H00C0C0C0&
      Height          =   4455
      Left            =   5040
      TabIndex        =   2
      Top             =   360
      Width           =   5895
      Begin VB.CommandButton cmdAjouterAlarme 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ajouter Alarme"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   3840
         Width           =   1455
      End
      Begin MSComCtl2.MonthView mvwChoixDate 
         Height          =   2820
         Left            =   1560
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   4974
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   0
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         StartOfWeek     =   152633345
         CurrentDate     =   37854
      End
      Begin VB.CommandButton cmdCopier 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Copier"
         Height          =   495
         Left            =   4560
         TabIndex        =   13
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton cmdAjouterCédule 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ajouter Cédule"
         Height          =   495
         Left            =   1680
         TabIndex        =   11
         Top             =   3840
         Width           =   1455
      End
      Begin VB.CommandButton cmdsupprimer 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Supprimer"
         Height          =   495
         Left            =   3240
         TabIndex        =   12
         Top             =   3840
         Width           =   1215
      End
      Begin MSComctlLib.ListView Lstjour 
         Height          =   3375
         Left            =   0
         TabIndex        =   8
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   5953
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "no"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "nom"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "début"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "fin"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "client"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Tansport"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Transport / Projet"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4560
         TabIndex        =   7
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Client"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3000
         TabIndex        =   6
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Fin"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1680
         TabIndex        =   5
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Début"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Nom"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame fraAlarme 
      Height          =   4455
      Left            =   5040
      TabIndex        =   32
      Top             =   360
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CommandButton cmdEnregistrerAlarme 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Enregistrer"
         Height          =   495
         Left            =   240
         TabIndex        =   37
         Top             =   3720
         Width           =   1455
      End
      Begin VB.CommandButton cmdAnnulerAlarme 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Annuler"
         Height          =   495
         Left            =   3600
         TabIndex        =   38
         Top             =   3720
         Width           =   1455
      End
      Begin VB.TextBox txtMessage 
         Height          =   1005
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   35
         Top             =   960
         Width           =   3855
      End
      Begin MSMask.MaskEdBox mskHeure 
         Height          =   255
         Left            =   1200
         TabIndex        =   34
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label9 
         Caption         =   "Message :"
         Height          =   255
         Left            =   360
         TabIndex        =   36
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Heure :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Frame frajour 
      Height          =   4455
      Left            =   5040
      TabIndex        =   14
      Top             =   360
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CommandButton cmdRafraichir 
         Caption         =   "Rafraîchir"
         Height          =   285
         Left            =   4680
         TabIndex        =   64
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CommandButton cmdRechercher 
         Caption         =   "Rechercher"
         Height          =   285
         Left            =   4680
         TabIndex        =   63
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox txtnoprojet 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1080
         TabIndex        =   24
         Top             =   1680
         Width           =   2175
      End
      Begin VB.CheckBox chkfin 
         Caption         =   "Fin Projet"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4080
         TabIndex        =   27
         Top             =   1680
         Width           =   855
      End
      Begin VB.ComboBox cmbclient 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1080
         TabIndex        =   28
         Top             =   2280
         Width           =   4695
      End
      Begin VB.CommandButton cmdtransport 
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
         Height          =   285
         Left            =   3480
         TabIndex        =   26
         Top             =   1680
         Width           =   375
      End
      Begin VB.ComboBox cmbtransport 
         Height          =   315
         ItemData        =   "frmcédule.frx":0442
         Left            =   1080
         List            =   "frmcédule.frx":044F
         Sorted          =   -1  'True
         TabIndex        =   25
         Text            =   "cmbtransport"
         Top             =   1680
         Width           =   2175
      End
      Begin VB.CommandButton cmdAnnulerCédule 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Annuler"
         Height          =   495
         Left            =   3240
         TabIndex        =   31
         Top             =   3840
         Width           =   1455
      End
      Begin VB.CommandButton cmdEnregistrerCédule 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Enregistrer"
         Height          =   495
         Left            =   120
         TabIndex        =   30
         Top             =   3840
         Width           =   1455
      End
      Begin VB.ComboBox cmbheurefin 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmcédule.frx":0467
         Left            =   2760
         List            =   "frmcédule.frx":04FB
         TabIndex        =   21
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ComboBox cmbheuredébut 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmcédule.frx":063B
         Left            =   1320
         List            =   "frmcédule.frx":06CF
         TabIndex        =   18
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ComboBox cmbemployé 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1080
         TabIndex        =   16
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label lblprojet 
         Caption         =   "No. Projet"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lbltransport 
         Caption         =   "Transport"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "à"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2520
         TabIndex        =   20
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "de"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1080
         TabIndex        =   19
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Client"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Cédulé"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Employé"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Fin des projets"
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
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "Mar"
      Height          =   255
      Left            =   5520
      TabIndex        =   62
      Top             =   5040
      Width           =   615
   End
End
Attribute VB_Name = "frmCédule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const I_LVW_JOUR_NO As Integer = 0
Private Const I_LVW_JOUR_NOM As Integer = 1
Private Const I_LVW_JOUR_DEBUT As Integer = 2
Private Const I_LVW_JOUR_FIN As Integer = 3
Private Const I_LVW_JOUR_CLIENT As Integer = 4
Private Const I_LVW_JOUR_TRANSPORT As Integer = 5

Private Const I_LVW_SEMAINE_NO As Integer = 0
Private Const I_LVW_SEMAINE_NOM As Integer = 1
Private Const I_LVW_SEMAINE_HEURE As Integer = 2

Private m_datDateChoisie As Date
Private m_bModeAjouter As Boolean
Private m_bMonthViewHasFocus As Boolean

Private Sub chkfin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 On Error GoTo Oups

 If m_bMonthViewHasFocus = True Then
 If chkfin.Value = vbChecked Then
 chkfin.Value = vbUnchecked
 Else
 chkfin.Value = vbChecked
 End If
 End If

 Exit Sub

Oups:

 wOups "frmCédule", "chkfin_MouseUp", Err, Err.number, Err.Description
End Sub

Private Sub cmdAjouterAlarme_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 On Error GoTo Oups

 If m_bMonthViewHasFocus = True Then
 Call cmdAjouterAlarme_Click
 End If

 Exit Sub

Oups:

 wOups "frmCédule", "cmdAjouterAlarme_MouseUp", Err, Err.number, Err.Description
End Sub

Private Sub cmdAjouterCédule_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 On Error GoTo Oups

 If m_bMonthViewHasFocus = True Then
 Call cmdAjouterCédule_Click
 End If

 Exit Sub

Oups:

 wOups "frmCédule", "cmdAjouterCédule_MouseUp", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnulerAlarme_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 On Error GoTo Oups

 If m_bMonthViewHasFocus = True Then
 Call cmdAnnulerAlarme_Click
 End If

 Exit Sub

Oups:

 wOups "frmCédule", "cmdAnnulerAlarme_MouseUp", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnulerCédule_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 On Error GoTo Oups

 If m_bMonthViewHasFocus = True Then
 Call cmdAnnulerCédule_Click
 End If

 Exit Sub

Oups:

 wOups "frmCédule", "cmdAnnulerCédule_MouseUp", Err, Err.number, Err.Description
End Sub

Private Sub cmdCopier_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 On Error GoTo Oups

 If m_bMonthViewHasFocus = True Then
 Call cmdCopier_Click
 End If

 Exit Sub

Oups:

 wOups "frmCédule", "cmdCopier_MouseUp", Err, Err.number, Err.Description
End Sub

Private Sub cmdEnregistrerAlarme_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 On Error GoTo Oups

 If m_bMonthViewHasFocus = True Then
 Call cmdEnregistrerAlarme_Click
 End If

 Exit Sub

Oups:

 wOups "frmCédule", "cmdEnregistrerAlarme_MouseUp", Err, Err.number, Err.Description
End Sub

Private Sub cmdEnregistrerCédule_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 On Error GoTo Oups

 If m_bMonthViewHasFocus = True Then
 Call cmdEnregistrerCédule_Click
 End If

 Exit Sub

Oups:

 wOups "frmCédule", "cmdEnregistrerCédule_MouseUp", Err, Err.number, Err.Description
End Sub

Private Sub cmdRafraichir_Click()
 
 On Error GoTo Oups
 
 Call RemplirListerClient

 m_bMonthViewHasFocus = False

 Exit Sub

Oups:

 wOups "frmCédule", "cmdRafraichir_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdRafraichir_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 On Error GoTo Oups

 If m_bMonthViewHasFocus = True Then
 Call cmdRafraichir_Click
 End If

 Exit Sub

Oups:

 wOups "frmCédule", "cmdRafraichir_MouseUp", Err, Err.number, Err.Description
End Sub

Private Sub cmdRechercher_Click()

 On Error GoTo Oups
 
 'Remplis combo employé
 Dim rstClient As ADODB.Recordset
 Dim sRecherche As String

 sRecherche = InputBox("Quel est le texte à rechercher ?")

 If sRecherche <> "" Then
 Set rstClient = New ADODB.Recordset

 Call rstClient.Open("SELECT * FROM GrbClient WHERE INSTR(1, NomClient,'" & sRecherche & "') > 0 AND Supprimé = False ORDER BY NomClient", g_connData, adOpenDynamic, adLockOptimistic)
 
 Call cmbclient.Clear
 
 'Rempli tant il y a des employé
 Do While Not rstClient.EOF
 Call cmbclient.AddItem(rstClient.Fields("NomClient"))
 
 Call rstClient.MoveNext
  Loop
 
  Call rstClient.Close
  Set rstClient = Nothing
  End If

  m_bMonthViewHasFocus = False

  Exit Sub

Oups:

  wOups "frmCédule", "cmdRechercher_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdRechercher_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 On Error GoTo Oups

 If m_bMonthViewHasFocus = True Then
 Call cmdRechercher_Click
 End If

 Exit Sub

Oups:

 wOups "frmCédule", "cmdRechercher_MouseUp", Err, Err.number, Err.Description
End Sub

Private Sub cmdsupprimer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 On Error GoTo Oups

 If m_bMonthViewHasFocus = True Then
 Call cmdsupprimer_Click
 End If

 Exit Sub

Oups:

 wOups "frmCédule", "cmdSupprimer_MouseUp", Err, Err.number, Err.Description
End Sub

Private Sub cmdtransport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 On Error GoTo Oups

 If m_bMonthViewHasFocus = True Then
 Call cmdTransport_Click
 End If

 Exit Sub

Oups:

 wOups "frmCédule", "cmdTransport_MouseUp", Err, Err.number, Err.Description
End Sub

Private Sub Form_Resize()

 On Error GoTo Oups

 Call frasemaine.Refresh

 Call frajour.Refresh

 Exit Sub

Oups:

 wOups "frmCédule", "Form_Resize", Err, Err.number, Err.Description
End Sub

Private Sub chkfin_Click()

 On Error GoTo Oups

 'On change l'affichage sur click
 'Fin projet ou transport
 If chkfin.Value = vbUnchecked Then
 Call AfficherTransport
 Else
 Call AfficherProjet
 End If

 m_bMonthViewHasFocus = False

 Exit Sub

Oups:

 wOups "frmCédule", "chkfin_Click", Err, Err.number, Err.Description
End Sub

Private Sub AfficherTransport()

 On Error GoTo Oups

 'Affichage transport
 lblprojet.Visible = False
 txtnoprojet.Visible = False
 cmbtransport.Visible = True
 cmdtransport.Visible = True
 lbltransport.Visible = True

 Exit Sub

Oups:

 wOups "frmCédule", "AfficherTransport", Err, Err.number, Err.Description
End Sub

Private Sub AfficherProjet()

 On Error GoTo Oups
 
 'Affichage fin de projet
 lblprojet.Visible = True
 txtnoprojet.Visible = True
 cmbtransport.Visible = False
 cmdtransport.Visible = False
 lbltransport.Visible = False

 Exit Sub

Oups:

 wOups "frmCédule", "AfficherProjet", Err, Err.number, Err.Description
End Sub

Private Sub cmdAjouterCédule_Click()

 On Error GoTo Oups

 Dim iCompteur As Integer

 'Met en mode ajouter et affiche champ pour entrer des données
 m_bModeAjouter = True

 fraliste.Visible = False
 fraAlarme.Visible = False
 frajour.Visible = True

 'Vide champ text
 cmbemployé.Text = vbNullString

 For iCompteur = 0 To cmbheuredébut.ListCount - 1
 If cmbheuredébut.LIST(iCompteur) = "8:00" Then
 cmbheuredébut.ListIndex = iCompteur

 Exit For
  End If
  Next

  For iCompteur = 0 To cmbheurefin.ListCount - 1
  If cmbheurefin.LIST(iCompteur) = "17:00" Then
  cmbheurefin.ListIndex = iCompteur

  Exit For
  End If
  Next

10 cmbtransport.Text = vbNullString
txtnoprojet.Text = vbNullString
cmbclient.Text = vbNullString
chkfin = vbUnchecked

m_bMonthViewHasFocus = False

Exit Sub

Oups:

wOups "frmCédule", "cmdAjouter_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdAjouterAlarme_Click()

 On Error GoTo Oups

 Dim iCompteur As Integer

 'Met en mode ajouter et affiche champ pour entrer des données
 m_bModeAjouter = True

 mskHeure.Text = ""
 txtMessage.Text = ""

 fraliste.Visible = False
 fraAlarme.Visible = True
 frajour.Visible = False

 mskHeure.Text = ""

 m_bMonthViewHasFocus = False

 Exit Sub

Oups:

  wOups "frmCédule", "cmdAjouter_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnulerAlarme_Click()

 On Error GoTo Oups

 'Quitte écran pour ajouter ou modifier
 fraliste.Visible = True
 fraAlarme.Visible = False
 frajour.Visible = False

 m_bMonthViewHasFocus = False

 Exit Sub

Oups:

 wOups "frmCédule", "cmdAnnulerAlarme_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnulerCédule_Click()

 On Error GoTo Oups

 'quitte ecran pour ajouté ou modifié
 fraliste.Visible = True
 fraAlarme.Visible = False
 frajour.Visible = False

 m_bMonthViewHasFocus = False

 Exit Sub

Oups:

 wOups "frmCédule", "cmdAnnuler_Click", Err, Err.number, Err.Description
End Sub

Private Sub CopierAlarme(ByVal datDate As Date)

 On Error GoTo Oups

 Dim sDate As String
 Dim rstAlarme As ADODB.Recordset
 Dim rstCopieAlarme As ADODB.Recordset
 Dim iCompteur As Integer
 
 sDate = ConvertDate(datDate)
 
 If sDate <> vbNullString Then
 datDate = DateSerial(Left$(sDate, 4), Mid$(sDate, 6, 2), Right$(sDate, 2))
 
 For iCompteur = 1 To Lstjour.ListItems.count
 If Lstjour.ListItems(iCompteur).Selected = True Then
 'ouvre la table
 Set rstAlarme = New ADODB.Recordset
  Set rstCopieAlarme = New ADODB.Recordset

  Call rstAlarme.Open("SELECT * FROM GrbAlarmes WHERE IDAlarme = " & Lstjour.ListItems(iCompteur).Text & " ORDER BY Initiale", g_connData, adOpenDynamic, adLockOptimistic)
  Call rstCopieAlarme.Open("SELECT * FROM GrbAlarmes", g_connData, adOpenDynamic, adLockOptimistic)
 
  Call rstCopieAlarme.AddNew
 
 'Ajoute l'enregistrement
  rstCopieAlarme.Fields("Initiale") = rstAlarme.Fields("Initiale")
  rstCopieAlarme.Fields("Date") = sDate
  rstCopieAlarme.Fields("Heure") = rstAlarme.Fields("Heure")
  rstCopieAlarme.Fields("JourSemaine") = Weekday(datDate)
 rstCopieAlarme.Fields("Type") = "C"
 
 Call rstCopieAlarme.Update
 
 'Quitte l'écran pour ajouté ou modifié
 fraliste.Visible = True
 fraAlarme.Visible = False
 frajour.Visible = False
 
 Call rstAlarme.Close
 Set rstAlarme = Nothing
 
 Call rstCopieAlarme.Close
 Set rstCopieAlarme = Nothing
 End If
 Next
 
 'Met à jour l'écran
 Call RemplirFinProjet
Call RemplirListerJour
 Call RemplirListerSemaine
 End If

Exit Sub

Oups:

 wOups "frmCédule", "CopierCedule", Err, Err.number, Err.Description
End Sub

Private Sub CopierCédule(ByVal datDate As Date)

 On Error GoTo Oups

 Dim sDate As String
 Dim rstCédule As ADODB.Recordset
 Dim rstCopieCédule As ADODB.Recordset
 Dim iCompteur As Integer
 
 sDate = ConvertDate(datDate)
 
 If sDate <> vbNullString Then
 datDate = DateSerial(Left$(sDate, 4), Mid$(sDate, 6, 2), Right$(sDate, 2))
 
 For iCompteur = 1 To Lstjour.ListItems.count
 If Lstjour.ListItems(iCompteur).Selected = True Then
 'ouvre la table
 Set rstCédule = New ADODB.Recordset
  Set rstCopieCédule = New ADODB.Recordset
 
  Call rstCédule.Open("SELECT * FROM Grbcédule WHERE noenreg = " & Lstjour.ListItems(iCompteur).Text & " ORDER BY initiale", g_connData, adOpenDynamic, adLockOptimistic)
  Call rstCopieCédule.Open("SELECT * FROM Grbcédule", g_connData, adOpenDynamic, adLockOptimistic)
 
  Call rstCopieCédule.AddNew
 
 ''''''''''''''''''''''''''
 'ajoute l'enregistrement
 ''''''''''''''''''''''''''
  rstCopieCédule.Fields("initiale") = rstCédule.Fields("initiale")
  rstCopieCédule.Fields("date_cedulé") = sDate
  rstCopieCédule.Fields("heure_début") = rstCédule.Fields("heure_début")
  rstCopieCédule.Fields("heure_fin") = rstCédule.Fields("heure_fin")
 rstCopieCédule.Fields("Client") = rstCédule.Fields("Client")
 rstCopieCédule.Fields("joursemaine") = Weekday(datDate)
 rstCopieCédule.Fields("finprojet") = rstCédule.Fields("finprojet")
 rstCopieCédule.Fields("transport") = rstCédule.Fields("transport")
 
 Call rstCopieCédule.Update
 
 'quitte l'écran pour ajouté ou modifié
 fraliste.Visible = True
 frajour.Visible = False
 
 Call rstCédule.Close
 Set rstCédule = Nothing

 Call rstCopieCédule.Close
 Set rstCopieCédule = Nothing
 End If
Next
 
 'met a jour l'écran
 Call RemplirFinProjet
 Call RemplirListerJour
 Call RemplirListerSemaine
 End If

Exit Sub

Oups:

 wOups "frmCédule", "CopierCedule", Err, Err.number, Err.Description
End Sub

Private Sub cmdCopier_Click()

 On Error GoTo Oups

 If Lstjour.ListItems.count > 0 Then
 mvwChoixDate.Month = mvwJanuary
 mvwChoixDate.Day = 1
 
 mvwChoixDate.Year = mvwSelection.Year
 mvwChoixDate.Month = mvwSelection.Month
 mvwChoixDate.Day = mvwSelection.Day
 
 mvwChoixDate.Visible = True
 
 Call mvwChoixDate.SetFocus
 End If

 m_bMonthViewHasFocus = False

  Exit Sub

Oups:

  wOups "frmCédule", "cmdCopier_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdEnregistrerAlarme_Click()

 On Error GoTo Oups

 'Enregistre
 Dim rstAlarme As ADODB.Recordset
 Dim rstEmploye As ADODB.Recordset
 Dim iNoEmploye As Integer

 If mskHeure.Text <> "" Then
 If IsDate(mskHeure.Text) Then
 If txtMessage.Text <> "" Then
 Set rstAlarme = New ADODB.Recordset
 Set rstEmploye = New ADODB.Recordset

 'Ouvre la table
 If m_bModeAjouter = True Then
 Call rstAlarme.Open("SELECT * FROM GrbAlarmes", g_connData, adOpenDynamic, adLockOptimistic)
 
  m_bModeAjouter = False

  Call rstAlarme.AddNew
  Else
  Call rstAlarme.Open("SELECT * FROM GrbAlarmes WHERE IDAlarme = " & Lstjour.SelectedItem.Text, g_connData, adOpenDynamic, adLockOptimistic)
  End If

  Call rstEmploye.Open("SELECT * FROM GrbEmployés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)

  iNoEmploye = rstEmploye.Fields("NoEmploye")
 
  Call rstEmploye.Close
 Set rstEmploye = Nothing
 
 'Ajoute l'enregistrement
 rstAlarme.Fields("NoEmploye") = iNoEmploye
 rstAlarme.Fields("Date") = ConvertDate(m_datDateChoisie)
 rstAlarme.Fields("Heure") = mskHeure.Text
 rstAlarme.Fields("Message") = txtMessage.Text
 rstAlarme.Fields("JourSemaine") = Weekday(m_datDateChoisie)
 rstAlarme.Fields("TypeCédule") = "C"
 
 'Quitte ecran pour ajouté ou modifié
 fraliste.Visible = True
 fraAlarme.Visible = False
 frajour.Visible = False
 
 Call rstAlarme.Update
 
 Call rstAlarme.Close
 Set rstAlarme = Nothing
 
 'Met à jour l'écran
 Call RemplirFinProjet
 Call RemplirListerJour
 Call RemplirListerSemaine
 Else
 Call MsgBox("Il n'y a pas de message à afficher!", vbOKOnly, "Erreur")
 End If
1  Else
 Call MsgBox("L'heure est invalide!", vbOKOnly, "Erreur")
 End If
Else
 Call MsgBox("L'heure est obligatoire!", vbOKOnly, "Erreur")
End If

m_bMonthViewHasFocus = False

Exit Sub

Oups:

wOups "frmCédule", "cmdEnregistrerAlarme_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdEnregistrerCédule_Click()

 On Error GoTo Oups

 'Enregistre
 Dim rstCédule As ADODB.Recordset
 Dim rstEmployé As ADODB.Recordset
 
 If cmbemployé.ListIndex <> -1 Then
 Set rstCédule = New ADODB.Recordset
 Set rstEmployé = New ADODB.Recordset

 'Ouvre la table
 If m_bModeAjouter = True Then
 Call rstCédule.Open("SELECT * FROM Grbcédule", g_connData, adOpenDynamic, adLockOptimistic)
 
 m_bModeAjouter = False
 
 Call rstCédule.AddNew
 Else
  Call rstCédule.Open("SELECT * FROM Grbcédule WHERE noenreg = " & Lstjour.ListItems(Lstjour.SelectedItem.Index).Text & " ORDER BY initiale", g_connData, adOpenDynamic, adLockOptimistic)
  End If
 
 'Ajoute l'enregistrement
  Call rstEmployé.Open("SELECT initiale FROM Grbemployés WHERE noemploye = " & cmbemployé.ItemData(cmbemployé.ListIndex), g_connData, adOpenDynamic, adLockOptimistic)
 
  rstCédule.Fields("initiale") = rstEmployé.Fields("initiale")
 
  Call rstEmployé.Close
  Set rstEmployé = Nothing
 
  rstCédule.Fields("date_cedulé") = ConvertDate(m_datDateChoisie)
 
  If cmbheuredébut.Text = vbNullString Then
 rstCédule.Fields("heure_début") = " "
 Else
 rstCédule.Fields("heure_début") = cmbheuredébut.Text
 End If
 
 If cmbheurefin.Text = vbNullString Then
 rstCédule.Fields("heure_fin") = " "
 Else
 rstCédule.Fields("heure_fin") = cmbheurefin.Text
 End If
 
 If cmbclient.Text = vbNullString Then
 rstCédule.Fields("CLIENT") = " "
Else
 rstCédule.Fields("CLIENT") = cmbclient.Text
 End If
 
 rstCédule.Fields("joursemaine") = Weekday(m_datDateChoisie)
 
 rstCédule.Fields("finprojet") = chkfin.Value
 
 'Enregistre le champ finprojet ou transport
 If chkfin.Value = vbUnchecked Then
 If cmbtransport.Text = vbNullString Then
1  rstCédule.Fields("transport") = " "
 Else
 rstCédule.Fields("transport") = cmbtransport.Text
 End If
 Else
 If txtnoprojet.Text = vbNullString Then
 rstCédule.Fields("transport") = " "
 Else
 rstCédule.Fields("transport") = txtnoprojet.Text
 End If
 End If
 
 Call rstCédule.Update
 
 'Quitte ecran pour ajouté ou modifié
 fraliste.Visible = True
fraAlarme.Visible = False
 frajour.Visible = False
 
Call rstCédule.Close
 Set rstCédule = Nothing
 
 'Met à jour l'écran
Call RemplirFinProjet
 Call RemplirListerJour
Call RemplirListerSemaine
Else
Call MsgBox("Aucun employé de sélectionné!")
End If

m_bMonthViewHasFocus = False

Exit Sub

Oups:

wOups "frmCédule", "cmdEnregistrerCédule_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdsupprimer_Click()

 On Error GoTo Oups
 
 'Supprime la cédule selectionnée
 Dim iCompteur As Integer
 
 If Lstjour.ListItems.count > 0 Then
 If MsgBox("Voulez-vous supprimer ce(ces) rendez-vous?", vbYesNo) = vbYes Then
 For iCompteur = 1 To Lstjour.ListItems.count
 If Lstjour.ListItems(iCompteur).Selected = True Then
 If Lstjour.ListItems(iCompteur).Tag = "A" Then
 Call g_connData.Execute("DELETE * FROM GrbAlarmes WHERE IDAlarme = " & Lstjour.ListItems(iCompteur).Text)
 Else
 Call g_connData.Execute("DELETE * FROM GrbCédule WHERE noenreg = " & Lstjour.ListItems(iCompteur).Text)
 End If
  End If
  Next
 
 'Mise à jour des ListViews
  Call RemplirFinProjet
  Call RemplirListerJour
  Call RemplirListerSemaine
  End If
  End If

  m_bMonthViewHasFocus = False

10 Exit Sub

Oups:

wOups "frmCédule", "cmdSupprimer_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdTransport_Click()

 On Error GoTo Oups

 Dim rstTransport As ADODB.Recordset
 Dim sTransport As String
 Dim iCompteur As Integer
 
 sTransport = cmbtransport.Text
 
 'Si'l y a un transport
 If cmbtransport.Text <> vbNullString Then
 'Si le transport existe, on demande si on veut le supprimer
 'sinon, on demande si on veut l'ajouter
 If ComboContient(cmbtransport, sTransport) Then
 'Si réponse oui pour supprimer
 If MsgBox("Voulez-vous supprimer le transport " & cmbtransport.Text & "?", vbYesNo) = vbYes Then
 Call g_connData.Execute("DELETE * FROM Grbtransport WHERE transport = '" & Replace(cmbtransport.Text, "'", "''") & "'")
 Else
 'Sinon demande si veut ajouter un nouveau transport
 If MsgBox("Voulez-vous ajouter un transport?", vbYesNo) = vbYes Then
  sTransport = InputBox("Veuillez entrer son nom!")
 
 'Si quelque chose d'entrer
  If sTransport <> vbNullString Then
  If Not ComboContient(cmbtransport, sTransport) Then
  Set rstTransport = New ADODB.Recordset

  Call rstTransport.Open("SELECT * FROM GrbTransport", g_connData, adOpenDynamic, adLockOptimistic)
 
  Call rstTransport.AddNew

  rstTransport.Fields("transport").Value = sTransport
 
  Call rstTransport.Update
 
 Call rstTransport.Close
 Set rstTransport = Nothing
 Else
 Call MsgBox("Ce transport existe déjà!")
 End If
 End If
 End If
 End If
 Else
 'Demande si veut ajouter un nouveau transport
 If MsgBox("Voulez-vous ajouter un transport?", vbYesNo) = vbYes Then
 sTransport = InputBox("Veuillez entrer son nom!")
 
 'Si quelque chose d'entrer
 If sTransport <> vbNullString Then
 If Not ComboContient(cmbtransport, sTransport) Then
 Set rstTransport = New ADODB.Recordset

 Call rstTransport.Open("SELECT * FROM Grbtransport", g_connData, adOpenDynamic, adLockOptimistic)

 Call rstTransport.AddNew
 rstTransport.Fields("transport") = sTransport
 
 Call rstTransport.Update
 
 Call rstTransport.Close
1  Set rstTransport = Nothing
 Else
 Call MsgBox("Ce transport existe déjà!")
 End If
 End If
 End If
 End If
Else
 'Demande si veut ajouter un nouveau transport
 If MsgBox("Voulez-vous ajouter un transport?", vbYesNo) = vbYes Then
 sTransport = InputBox("Veuillez entrer son nom!")
 
 'Si quelque chose d'entrer
 If sTransport <> vbNullString Then
 If Not ComboContient(cmbtransport, sTransport) Then
 Set rstTransport = New ADODB.Recordset

 Call rstTransport.Open("SELECT * FROM Grbtransport", g_connData, adOpenDynamic, adLockOptimistic)
 
 Call rstTransport.AddNew

 rstTransport.Fields("transport") = sTransport
 
 Call rstTransport.Update
 
 Call rstTransport.Close
 Set rstTransport = Nothing
 Else
 Call MsgBox("Ce transport existe déjà!")
 End If
End If
 End If
End If
 
 'Remplis combo transport
Call RemplirTransport
 
For iCompteur = 0 To cmbtransport.ListCount - 1
 If cmbtransport.LIST(iCompteur) = sTransport Then
 cmbtransport.ListIndex = iCompteur
 
 Exit For
 End If
Next
 
If cmbtransport.ListIndex = -1 Then
cmbtransport.ListIndex = 0
End If

3  m_bMonthViewHasFocus = False

Exit Sub

Oups:

3  wOups "frmCédule", "cmdtransport_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups

 Dim iCompteur As Integer

 mvwSelection.StartOfWeek = mvwSunday

 g_bCeduleOuverte = True

 'Met à jour l'écran
 mvwSelection.Year = Year(Date)
 mvwSelection.Month = Month(Date)
 mvwSelection.Day = Day(Date)

 m_datDateChoisie = Date
 
 'Rempli les combos
 Call RemplirListerEmployé
 Call RemplirTransport
 Call RemplirListerClient
  Call RemplirFinProjet
 
 'Rempli les ListViews
  Call RemplirListerJour
  Call RemplirListerSemaine
 
 'Sélectionne le jour de la semaine
  For iCompteur = 1 To 7
  If lstjoursemaine(iCompteur).Tag = m_datDateChoisie Then
  lstjoursemaine(iCompteur).BackColor = &HE0E0E0
  Else
  lstjoursemaine(iCompteur).BackColor = &HFFFFFF
End If
Next
 
Call AfficherTransport

Screen.MousePointer = vbDefault

Exit Sub

Oups:

wOups "frmCédule", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub RemplirFinProjet()

 On Error GoTo Oups
 
 'Remplis le combo transport
 Dim rstCedule As ADODB.Recordset
 
 Set rstCedule = New ADODB.Recordset
 
 Call rstCedule.Open("SELECT date_cedulé, transport FROM Grbcédule WHERE finprojet = 1 ORDER BY date_cedulé", g_connData, adOpenDynamic, adLockOptimistic)
 
 Call cmbfinprojet.Clear
 
 'Rempli tant il y a des employé
 Do While Not rstCedule.EOF
 DoEvents
 
 Call cmbfinprojet.AddItem(Trim$(CStr(rstCedule!transport)) & " " & ConvertDate(rstCedule!date_cedulé))
 
 Call rstCedule.MoveNext
 Loop
 
 Call rstCedule.Close
  Set rstCedule = Nothing

 'S'il y a des enregistrements, on sélectionne le premier
  If cmbfinprojet.ListCount > 0 Then
  cmbfinprojet.ListIndex = 0
  End If

  Exit Sub

Oups:

  wOups "frmCédule", "RemplirFinProjet", Err, Err.number, Err.Description
End Sub

Private Sub RemplirTransport()

 On Error GoTo Oups
 
 ''''''''''''''''''''''''
 'remplis combo transport
 ''''''''''''''''''''''''
 Dim rstTransport As ADODB.Recordset
 
 Set rstTransport = New ADODB.Recordset
 
 Call rstTransport.Open("SELECT * FROM Grbtransport ORDER BY transport", g_connData, adOpenDynamic, adLockOptimistic)
 
 Call cmbtransport.Clear
 
 'rempli tant il y a des employé
 Do While Not rstTransport.EOF
 Call cmbtransport.AddItem(rstTransport!transport)
 
 Call rstTransport.MoveNext
 Loop
 
 Call rstTransport.Close
 Set rstTransport = Nothing

  Exit Sub

Oups:

  wOups "frmCédule", "RemplirTransport", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListerEmployé()

 On Error GoTo Oups
 
 '''''''''''''''''''''''''
 ' Remplis combo employé '
 '''''''''''''''''''''''''
 Dim rstEmploye As ADODB.Recordset

 Set rstEmploye = New ADODB.Recordset

 Call rstEmploye.Open("SELECT * FROM Grbemployés WHERE Actif = True ORDER BY employe", g_connData, adOpenDynamic, adLockOptimistic)
 
 Call cmbemployé.Clear
 
 'rempli tant il y a des employé
 Do While Not rstEmploye.EOF
 Call cmbemployé.AddItem(rstEmploye!employe)
 cmbemployé.ItemData(cmbemployé.newIndex) = rstEmploye!noEmploye
 
 Call rstEmploye.MoveNext
 Loop
 
 Call rstEmploye.Close
  Set rstEmploye = Nothing

  Exit Sub

Oups:

  wOups "frmCédule", "RemplirListerEmployé", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListerClient()

 On Error GoTo Oups
 
 'remplis combo employé
 Dim rstClient As ADODB.Recordset

 Set rstClient = New ADODB.Recordset

 Call rstClient.Open("SELECT * FROM GrbClient WHERE Supprimé = False ORDER BY NomClient", g_connData, adOpenDynamic, adLockOptimistic)
 
 Call cmbclient.Clear
 
 'rempli tant il y a des employé
 Do While Not rstClient.EOF
 Call cmbclient.AddItem(rstClient!nomclient)
 
 Call rstClient.MoveNext
 Loop
 
 Call rstClient.Close
 Set rstClient = Nothing

  Exit Sub

Oups:

  wOups "frmCédule", "RemplirListerClient", Err, Err.number, Err.Description
End Sub

Public Sub RemplirListerJour()

 On Error GoTo Oups

 'Remplis lister une journée
 Dim rstCédule As ADODB.Recordset
 Dim itmCedule As ListItem
 
 'Vide le lister
 Call Lstjour.ListItems.Clear
 
 Set rstCédule = New ADODB.Recordset
 
 Call rstCédule.Open("SELECT * FROM Grbcédule WHERE date_cedulé = '" & ConvertDate(m_datDateChoisie) & "' ORDER BY initiale", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Tant il y a de employé cedulé , ajoute dans lister
 Do While Not rstCédule.EOF
 Set itmCedule = Lstjour.ListItems.Add
 
 itmCedule.Text = rstCédule!noenreg

 itmCedule.Tag = "C"

 If Not IsNull(rstCédule.Fields("Initiale")) Then
  itmCedule.SubItems(I_LVW_JOUR_NOM) = rstCédule!Initiale
  Else
  itmCedule.SubItems(I_LVW_JOUR_NOM) = ""
  End If
 
  If Not IsNull(rstCédule!heure_début) Then
  itmCedule.SubItems(I_LVW_JOUR_DEBUT) = rstCédule!heure_début
  Else
  itmCedule.SubItems(I_LVW_JOUR_DEBUT) = ""
End If
 
1 If Not IsNull(rstCédule!heure_fin) Then
 itmCedule.SubItems(I_LVW_JOUR_FIN) = rstCédule!heure_fin
 Else
 itmCedule.SubItems(I_LVW_JOUR_FIN) = ""
 End If
 
 If Not IsNull(rstCédule!CLIENT) Then
 itmCedule.SubItems(I_LVW_JOUR_CLIENT) = rstCédule!CLIENT
 Else
 itmCedule.SubItems(I_LVW_JOUR_CLIENT) = ""
 End If
 
 'si fin de projet marque numero de projet sinon transport
 If rstCédule!finprojet = 0 Then
 If Not IsNull(rstCédule!transport) Then
 itmCedule.SubItems(I_LVW_JOUR_TRANSPORT) = rstCédule!transport
 Else
 itmCedule.SubItems(I_LVW_JOUR_TRANSPORT) = ""
 End If
 
 'met en rouge ou en noir dépendant si fin de projet
 itmCedule.ListSubItems(I_LVW_JOUR_NOM).ForeColor = COLOR_NOIR
 itmCedule.ListSubItems(I_LVW_JOUR_DEBUT).ForeColor = COLOR_NOIR
1  itmCedule.ListSubItems(I_LVW_JOUR_FIN).ForeColor = COLOR_NOIR
 itmCedule.ListSubItems(I_LVW_JOUR_CLIENT).ForeColor = COLOR_NOIR
 itmCedule.ListSubItems(I_LVW_JOUR_TRANSPORT).ForeColor = COLOR_NOIR
 Else
 If Not IsNull(rstCédule!transport) Then
 itmCedule.SubItems(I_LVW_JOUR_TRANSPORT) = "Fin " + rstCédule!transport
 Else
 itmCedule.SubItems(I_LVW_JOUR_TRANSPORT) = "Fin"
 End If
 
 'Met en rouge ou en noir dépendant si fin de projet
 itmCedule.ListSubItems(I_LVW_JOUR_NOM).ForeColor = COLOR_ROUGE
 itmCedule.ListSubItems(I_LVW_JOUR_DEBUT).ForeColor = COLOR_ROUGE
 itmCedule.ListSubItems(I_LVW_JOUR_FIN).ForeColor = COLOR_ROUGE
 itmCedule.ListSubItems(I_LVW_JOUR_CLIENT).ForeColor = COLOR_ROUGE
 itmCedule.ListSubItems(I_LVW_JOUR_TRANSPORT).ForeColor = COLOR_ROUGE
 End If
 
Call rstCédule.MoveNext
Loop
 
2  Call rstCédule.Close
Set rstCédule = Nothing

2  Call RemplirListerJourAlarme

Exit Sub

Oups:

30 wOups "frmCédule", "RemplirListerJour", Err, Err.number, Err.Description
End Sub

Public Sub RemplirListerSemaine()

 On Error GoTo Oups
 
 'Remplis une semaine
 Dim rstCédule As ADODB.Recordset
 Dim iJourSemaine As Integer
 Dim datPremiereDate As Date
 Dim datDerniereDate As Date
 Dim iCompteur As Integer
 Dim sHeureDebutFin As String
 Dim itmSemaine As ListItem

 For iCompteur = 1 To 7
 'couleur par defaut entete de date
 lbljour(iCompteur - 1).ForeColor = vbWhite
 lbljourstr(iCompteur - 1).ForeColor = vbWhite

  Call lstjoursemaine(iCompteur).ListItems.Clear
  Next
 
  iJourSemaine = Weekday(m_datDateChoisie)
  datPremiereDate = m_datDateChoisie
  datDerniereDate = m_datDateChoisie
 
 'trouve premiere date de la semaine
  Do While Not Weekday(datPremiereDate) = 1
  datPremiereDate = datPremiereDate - 1
  Loop
 
 'trouve derniere date de la semaine
10 Do While Not Weekday(datDerniereDate) = 7
1 datDerniereDate = datDerniereDate + 1
Loop
 
 'selectionne la semaine courante
Set rstCédule = New ADODB.Recordset
 
Call rstCédule.Open("SELECT * FROM Grbcédule WHERE cdate(date_cedulé) <= cdate('" & CStr(datDerniereDate) & "') AND cdate(date_cedulé) >= cdate('" & CStr(datPremiereDate) & "') ORDER BY initiale", g_connData, adOpenDynamic, adLockOptimistic)
 
For iCompteur = 1 To 7
 'pour ecrire le jour
 lbljour(iCompteur - 1).Caption = Day(datPremiereDate + iCompteur - 1)
 
 'garde en memoire la date des lister
 lstjoursemaine(iCompteur).Tag = datPremiereDate + iCompteur - 1
Next
 
Do While Not rstCédule.EOF
 'ajoute dans le lister, dépendant le jour de la semaine
 Set itmSemaine = lstjoursemaine(rstCédule!joursemaine).ListItems.Add
 
 itmSemaine.Text = rstCédule!noenreg
 
 'si fin de projet marque numero de projet sinon transport
If rstCédule!finprojet = 0 Then
 If Not IsNull(rstCédule.Fields("Initiale")) Then
 itmSemaine.SubItems(I_LVW_SEMAINE_NOM) = rstCédule!Initiale
 Else
 itmSemaine.SubItems(I_LVW_SEMAINE_NOM) = ""
 End If
 
 If Not IsNull(rstCédule!heure_début) Then
1  sHeureDebutFin = Trim(rstCédule!heure_début + "-")
 Else
 sHeureDebutFin = "-"
 End If
 
 If Not IsNull(rstCédule!heure_fin) Then
 sHeureDebutFin = sHeureDebutFin + rstCédule!heure_fin
 Else
 sHeureDebutFin = sHeureDebutFin + " "
 End If
 
 itmSemaine.SubItems(I_LVW_SEMAINE_HEURE) = sHeureDebutFin
 
 'met en noir
 itmSemaine.ListSubItems(I_LVW_SEMAINE_NOM).ForeColor = COLOR_NOIR
 itmSemaine.ListSubItems(I_LVW_SEMAINE_HEURE).ForeColor = COLOR_NOIR
 Else
 itmSemaine.SubItems(I_LVW_SEMAINE_NOM) = "Fin"
 
 lbljour(rstCédule!joursemaine - 1).ForeColor = COLOR_ROUGE
 lbljourstr(rstCédule!joursemaine - 1).ForeColor = COLOR_ROUGE
 
 If Not IsNull(rstCédule!transport) Then
 sHeureDebutFin = rstCédule!transport
 Else
 sHeureDebutFin = " "
 End If
 
 itmSemaine.SubItems(I_LVW_SEMAINE_HEURE) = sHeureDebutFin
 
 'met en rouge
itmSemaine.ListSubItems(I_LVW_SEMAINE_NOM).ForeColor = COLOR_ROUGE
 itmSemaine.ListSubItems(I_LVW_SEMAINE_HEURE).ForeColor = COLOR_ROUGE
 End If
 
 Call rstCédule.MoveNext
Loop
 
Call rstCédule.Close
Set rstCédule = Nothing

Call RemplirListerSemaineAlarme

Exit Sub

Oups:

wOups "frmCédule", "RemplirListerSemaine", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListerJourAlarme()

 On Error GoTo Oups

 'Remplis lister une journée
 Dim rstAlarme As ADODB.Recordset
 Dim rstEmploye As ADODB.Recordset
 Dim iNoEmploye As Integer
 Dim itmAlarme As ListItem

 Set rstEmploye = New ADODB.Recordset

 Call rstEmploye.Open("SELECT * FROM GrbEmployés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 iNoEmploye = rstEmploye.Fields("NoEmploye")

 Call rstEmploye.Close
 Set rstEmploye = Nothing

 Set rstAlarme = New ADODB.Recordset

  Call rstAlarme.Open("SELECT * FROM GrbAlarmes WHERE Date = '" & ConvertDate(m_datDateChoisie) & "' AND NoEmploye = " & iNoEmploye & " AND TypeCédule = 'C' ORDER BY Date, Heure", g_connData, adOpenDynamic, adLockOptimistic)
 
 'Tant il y a de employé cedulé , ajoute dans lister
  Do While Not rstAlarme.EOF
  Set itmAlarme = Lstjour.ListItems.Add
 
  itmAlarme.Text = (rstAlarme.Fields("IDAlarme"))
  itmAlarme.Tag = "A"

  itmAlarme.SubItems(I_LVW_JOUR_NOM) = g_sInitiale
 
  If Not IsNull(rstAlarme.Fields("Heure")) Then
  itmAlarme.SubItems(I_LVW_JOUR_DEBUT) = rstAlarme.Fields("Heure")
Else
itmAlarme.SubItems(I_LVW_JOUR_DEBUT) = ""
 End If
 
 'Met en rouge ou en noir dépendant si fin de projet
 itmAlarme.ListSubItems(I_LVW_JOUR_NOM).ForeColor = COLOR_BLEU
 itmAlarme.ListSubItems(I_LVW_JOUR_DEBUT).ForeColor = COLOR_BLEU
 
 Call rstAlarme.MoveNext
Loop
 
Call rstAlarme.Close
Set rstAlarme = Nothing

Exit Sub

Oups:

wOups "frmCédule", "RemplirListerJourAlarme", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListerSemaineAlarme()

 On Error GoTo Oups

 'Remplis une semaine
 Dim rstAlarme As ADODB.Recordset
 Dim rstEmploye As ADODB.Recordset
 Dim iNoEmploye As Integer
 Dim iJourSemaine As Integer
 Dim datPremiereDate As Date
 Dim datDerniereDate As Date
 Dim iCompteur As Integer
 Dim itmSemaine As ListItem
 
 iJourSemaine = Weekday(m_datDateChoisie)
 datPremiereDate = m_datDateChoisie
  datDerniereDate = m_datDateChoisie
 
 'Trouve première date de la semaine
  Do While Not Weekday(datPremiereDate) = 1
  datPremiereDate = datPremiereDate - 1
  Loop
 
 'Trouve dernière date de la semaine
  Do While Not Weekday(datDerniereDate) = 7
  datDerniereDate = datDerniereDate + 1
  Loop

  Set rstEmploye = New ADODB.Recordset

10 Call rstEmploye.Open("SELECT * FROM GrbEmployés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)

iNoEmploye = rstEmploye.Fields("NoEmploye")

Call rstEmploye.Close
Set rstEmploye = Nothing
 
 'Sélectionne la semaine courante
Set rstAlarme = New ADODB.Recordset
 
Call rstAlarme.Open("SELECT * FROM GrbAlarmes WHERE Date BETWEEN '" & ConvertDate(datPremiereDate) & "' AND '" & ConvertDate(datDerniereDate) & "' AND NoEmploye = " & iNoEmploye & " AND TypeCédule = 'C' ORDER BY Date, Heure", g_connData, adOpenDynamic, adLockOptimistic)
 
 'iSemaine est le numero du lister, jour de semaine
For iCompteur = 1 To 7
 'pour ecrire le jour
 lbljour(iCompteur - 1).Caption = Day(datPremiereDate + iCompteur - 1)
 
 'garde en memoire la date des lister
 lstjoursemaine(iCompteur).Tag = datPremiereDate + iCompteur - 1
Next
 
Do While Not rstAlarme.EOF
 'ajoute dans le lister, dépendant le jour de la semaine
 Set itmSemaine = lstjoursemaine(rstAlarme.Fields("JourSemaine")).ListItems.Add
 
itmSemaine.Text = rstAlarme.Fields("IDAlarme")
 
 'si fin de projet marque numero de projet sinon transport
 itmSemaine.SubItems(I_LVW_SEMAINE_NOM) = g_sInitiale
 
 itmSemaine.SubItems(I_LVW_SEMAINE_HEURE) = rstAlarme.Fields("Heure")
 
 'met en noir
 itmSemaine.ListSubItems(I_LVW_SEMAINE_NOM).ForeColor = COLOR_BLEU
 itmSemaine.ListSubItems(I_LVW_SEMAINE_HEURE).ForeColor = COLOR_BLEU
 
 Call rstAlarme.MoveNext
 Loop
 
1  Call rstAlarme.Close
 Set rstAlarme = Nothing

 Exit Sub

Oups:

wOups "frmCédule", "RemplirListerSemaineAlarme", Err, Err.number, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)

 On Error GoTo Oups

 g_bCeduleOuverte = False

 Exit Sub

Oups:

 wOups "frmCédule", "Form_Unload", Err, Err.number, Err.Description
End Sub

Private Sub Lstjour_DblClick()

 On Error GoTo Oups

 Dim rstCédule As ADODB.Recordset
 Dim rstAlarme As ADODB.Recordset
 Dim rstEmployé As ADODB.Recordset
 Dim iCompteur As Integer

 'Affiche en mode modification
 m_bModeAjouter = False
 
 If Lstjour.ListItems.count > 0 Then
 fraliste.Visible = False

 If Lstjour.SelectedItem.Tag = "C" Then
 frajour.Visible = True
 fraAlarme.Visible = False

 'Ouvre la table
  Set rstCédule = New ADODB.Recordset
  Set rstEmployé = New ADODB.Recordset
 
  Call rstCédule.Open("SELECT * FROM Grbcédule WHERE noenreg = " & Lstjour.SelectedItem.Text & " ORDER BY initiale", g_connData, adOpenDynamic, adLockOptimistic)
  Call rstEmployé.Open("SELECT * FROM Grbemployés WHERE initiale = '" & rstCédule.Fields("initiale") & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
 'si il y a employé, affiche a l'écran pour modification
  If Not rstEmployé.EOF Then
  For iCompteur = 0 To cmbemployé.ListCount - 1
  If cmbemployé.LIST(iCompteur) = rstEmployé.Fields("Employe") Then
  cmbemployé.ListIndex = iCompteur

 Exit For
 End If
 Next

 cmbheuredébut.Text = rstCédule.Fields("heure_début")
 cmbheurefin.Text = rstCédule.Fields("heure_fin")
 
 If IsNull(rstCédule!CLIENT) Then
 cmbclient.Text = " "
 Else
 cmbclient.Text = rstCédule!CLIENT
 End If
 
 chkfin.Value = rstCédule!finprojet
 
 If IsNull(rstCédule!transport) Then
 cmbtransport.Text = " "
 Else
 cmbtransport.Text = rstCédule.Fields("transport")
 End If
 
 If IsNull(rstCédule!transport) Then
 txtnoprojet.Text = " "
 Else
1  txtnoprojet.Text = rstCédule!transport
 End If
 
 'Affiche fin de projet ou transport
 If chkfin = vbUnchecked Then
 Call AfficherTransport
 Else
 Call AfficherProjet
 End If
 End If
 
 Call rstCédule.Close
 Set rstCédule = Nothing
 
 Call rstEmployé.Close
 Set rstEmployé = Nothing
 Else
 frajour.Visible = False
 fraAlarme.Visible = True

 mskHeure.Text = ""
 txtMessage.Text = ""

 'Ouvre la table
 Set rstAlarme = New ADODB.Recordset

 Call rstAlarme.Open("SELECT * FROM GrbAlarmes WHERE IDAlarme = " & Lstjour.SelectedItem.Text & " ORDER BY NoEmploye", g_connData, adOpenDynamic, adLockOptimistic)
 
 mskHeure.Text = rstAlarme.Fields("Heure")
 
 txtMessage.Text = rstAlarme.Fields("Message")
 
 Call rstAlarme.Close
Set rstAlarme = Nothing
 End If
Else
 fraliste.Visible = True
 frajour.Visible = False
End If

Exit Sub

Oups:

wOups "frmCédule", "Lstjour_DblClick", Err, Err.number, Err.Description
End Sub

Private Sub Lstjour_KeyDown(KeyCode As Integer, Shift As Integer)

 On Error GoTo Oups

 If Lstjour.ListItems.count > 0 Then
 If KeyCode = vbKeyDelete Then
 If MsgBox("Voulez-vous supprimer cette enregistrement?", vbYesNo) = vbYes Then
 If Lstjour.SelectedItem.Tag = "A" Then
 Call g_connData.Execute("DELETE * FROM GrbAlarmes WHERE IDAlarme = " & Lstjour.SelectedItem.Text)
 Else
 Call g_connData.Execute("DELETE * FROM GrbCédule WHERE noenreg = " & Lstjour.SelectedItem.Text)
 End If

 'Mise à jour des lister
 Call RemplirFinProjet
 Call RemplirListerJour
  Call RemplirListerSemaine
  End If
  End If
  End If

  Exit Sub

Oups:

  wOups "frmCédule", "Lstjour_KeyDown", Err, Err.number, Err.Description
End Sub

Private Sub lstjoursemaine_Click(Index As Integer)

 On Error GoTo Oups

 Dim iCompteur As Integer
 Dim sDate As String
 Dim iNbreJour As Integer
 
 'Initialise la couleur en blanc
 For iCompteur = 1 To 7
 lstjoursemaine(iCompteur).BackColor = &HFFFFFF
 Next
 
 'Sélectionne jour de semaine
 lstjoursemaine(Index).BackColor = &HE0E0E0

 sDate = lstjoursemaine(Index).Tag

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
Call RemplirListerJour

 fraliste.Visible = True
frajour.Visible = False

 Call Lstjour.SetFocus

1  Exit Sub

Oups:

 wOups "frmCédule", "lstjoursemaine_Click", Err, Err.number, Err.Description
End Sub

Private Sub mvwChoixDate_DateClick(ByVal DateClicked As Date)

 On Error GoTo Oups

 If Lstjour.SelectedItem.Tag = "A" Then
 Call CopierAlarme(DateClicked)
 Else
 Call CopierCédule(DateClicked)
 End If
 
 mvwChoixDate.Visible = False

 Exit Sub

Oups:

 wOups "frmCédule", "mvwChoixDate_DateClick", Err, Err.number, Err.Description
End Sub

Private Sub mvwChoixDate_LostFocus()

 On Error GoTo Oups
 
 mvwChoixDate.Visible = False

 Exit Sub

Oups:

 wOups "frmCédule", "mvwChoixDate_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskHeure_GotFocus()

 On Error GoTo Oups
 
 'Format d'heure
 mskHeure.mask = "##:##"

 Exit Sub

Oups:

 wOups "frmCédule", "mskHeure_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mskHeure_LostFocus()

 On Error GoTo Oups

 'Enlève le mask
 mskHeure.mask = vbNullString
 
 'Vide le champs si l'utilisateur n'a rien écrit
 If mskHeure.Text = "__:__" Then
 mskHeure.Text = vbNullString
 End If

 Exit Sub

Oups:

 wOups "frmCédule", "mskHeure_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub AfficherDate()

 On Error GoTo Oups

 'Affiche horaire de la journée et de la semaine
 'dépendant la sélection dans le calendrier
 Dim iCompteur As Integer
 
 'Date choisie
 m_datDateChoisie = DateSerial(mvwSelection.Year, mvwSelection.Month, mvwSelection.Day)

 'Affiche horaire jour et semaine
 Call RemplirListerJour
 Call RemplirListerSemaine

 'Sélectionne jour de la semaine
 For iCompteur = 1 To 7
 If lstjoursemaine(iCompteur).Tag = m_datDateChoisie Then
 lstjoursemaine(iCompteur).BackColor = &HE0E0E0
 Else
 lstjoursemaine(iCompteur).BackColor = &HFFFFFF
 End If
  Next

 'Affiche cédule une journée
  fraliste.Visible = True
  fraAlarme.Visible = False
  frajour.Visible = False

  Exit Sub

Oups:

  wOups "frmCédule", "AfficherDate", Err, Err.number, Err.Description
End Sub

Private Sub mvwSelection_GotFocus()

 On Error GoTo Oups

 m_bMonthViewHasFocus = True

 Exit Sub

Oups:

 wOups "frmCédule", "mvwSelection_GotFocus", Err, Err.number, Err.Description
End Sub

Private Sub mvwSelection_SelChange(ByVal StartDate As Date, ByVal EndDate As Date, Cancel As Boolean)

 On Error GoTo Oups

 If Month(m_datDateChoisie) <> mvwSelection.Month Or _
 Year(m_datDateChoisie) <> mvwSelection.Year Or _
 Day(m_datDateChoisie) <> mvwSelection.Day Then
 Call AfficherDate
 End If

 Exit Sub

Oups:

 wOups "frmCédule", "mvwSelection_SelChange", Err, Err.number, Err.Description
End Sub
