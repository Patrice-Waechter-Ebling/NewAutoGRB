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
      StartOfWeek     =   90243073
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
         StartOfWeek     =   90243073
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

Private Const I_LVW_JOUR_NO        As Integer = 0
Private Const I_LVW_JOUR_NOM       As Integer = 1
Private Const I_LVW_JOUR_DEBUT     As Integer = 2
Private Const I_LVW_JOUR_FIN       As Integer = 3
Private Const I_LVW_JOUR_CLIENT    As Integer = 4
Private Const I_LVW_JOUR_TRANSPORT As Integer = 5

Private Const I_LVW_SEMAINE_NO     As Integer = 0
Private Const I_LVW_SEMAINE_NOM    As Integer = 1
Private Const I_LVW_SEMAINE_HEURE  As Integer = 2

Private m_datDateChoisie     As Date
Private m_bModeAjouter       As Boolean
Private m_bMonthViewHasFocus As Boolean

Private Sub chkfin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5       On Error GoTo AfficherErreur

10      If m_bMonthViewHasFocus = True Then
15        If chkfin.Value = vbChecked Then
20          chkfin.Value = vbUnchecked
25        Else
30          chkfin.Value = vbChecked
35        End If
40      End If

45      Exit Sub

AfficherErreur:

50      woups "frmCédule", "chkfin_MouseUp", Err, Erl
End Sub

Private Sub cmdAjouterAlarme_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5       On Error GoTo AfficherErreur

10      If m_bMonthViewHasFocus = True Then
15        Call cmdAjouterAlarme_Click
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmCédule", "cmdAjouterAlarme_MouseUp", Err, Erl
End Sub

Private Sub cmdAjouterCédule_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5       On Error GoTo AfficherErreur

10      If m_bMonthViewHasFocus = True Then
15        Call cmdAjouterCédule_Click
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmCédule", "cmdAjouterCédule_MouseUp", Err, Erl
End Sub

Private Sub cmdAnnulerAlarme_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5       On Error GoTo AfficherErreur

10      If m_bMonthViewHasFocus = True Then
15        Call cmdAnnulerAlarme_Click
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmCédule", "cmdAnnulerAlarme_MouseUp", Err, Erl
End Sub

Private Sub cmdAnnulerCédule_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5       On Error GoTo AfficherErreur

10      If m_bMonthViewHasFocus = True Then
15        Call cmdAnnulerCédule_Click
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmCédule", "cmdAnnulerCédule_MouseUp", Err, Erl
End Sub

Private Sub cmdCopier_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5       On Error GoTo AfficherErreur

10      If m_bMonthViewHasFocus = True Then
15        Call cmdCopier_Click
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmCédule", "cmdCopier_MouseUp", Err, Erl
End Sub

Private Sub cmdEnregistrerAlarme_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5       On Error GoTo AfficherErreur

10      If m_bMonthViewHasFocus = True Then
15        Call cmdEnregistrerAlarme_Click
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmCédule", "cmdEnregistrerAlarme_MouseUp", Err, Erl
End Sub

Private Sub cmdEnregistrerCédule_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5       On Error GoTo AfficherErreur

10      If m_bMonthViewHasFocus = True Then
15        Call cmdEnregistrerCédule_Click
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmCédule", "cmdEnregistrerCédule_MouseUp", Err, Erl
End Sub

Private Sub cmdRafraichir_Click()
  
5       On Error GoTo AfficherErreur
  
10      Call RemplirListerClient

15      m_bMonthViewHasFocus = False

20      Exit Sub

AfficherErreur:

25      woups "frmCédule", "cmdRafraichir_Click", Err, Erl
End Sub

Private Sub cmdRafraichir_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5       On Error GoTo AfficherErreur

10      If m_bMonthViewHasFocus = True Then
15        Call cmdRafraichir_Click
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmCédule", "cmdRafraichir_MouseUp", Err, Erl
End Sub

Private Sub cmdRechercher_Click()

5       On Error GoTo AfficherErreur
        
        'Remplis combo employé
10      Dim rstClient  As ADODB.Recordset
15      Dim sRecherche As String

20      sRecherche = InputBox("Quel est le texte à rechercher ?")

25      If sRecherche <> "" Then
30        Set rstClient = New ADODB.Recordset

35        Call rstClient.Open("SELECT * FROM GRB_Client WHERE INSTR(1, NomClient,'" & sRecherche & "') > 0 AND Supprimé = False ORDER BY NomClient", g_connData, adOpenDynamic, adLockOptimistic)
    
40        Call cmbclient.Clear
    
          'Rempli tant il y a des employé
45        Do While Not rstClient.EOF
50          Call cmbclient.AddItem(rstClient.Fields("NomClient"))
        
55          Call rstClient.MoveNext
60        Loop
        
65        Call rstClient.Close
70        Set rstClient = Nothing
75      End If

80      m_bMonthViewHasFocus = False

85      Exit Sub

AfficherErreur:

90      woups "frmCédule", "cmdRechercher_Click", Err, Erl
End Sub

Private Sub cmdRechercher_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5       On Error GoTo AfficherErreur

10      If m_bMonthViewHasFocus = True Then
15        Call cmdRechercher_Click
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmCédule", "cmdRechercher_MouseUp", Err, Erl
End Sub

Private Sub cmdsupprimer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5       On Error GoTo AfficherErreur

10      If m_bMonthViewHasFocus = True Then
15        Call cmdsupprimer_Click
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmCédule", "cmdSupprimer_MouseUp", Err, Erl
End Sub

Private Sub cmdtransport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

5       On Error GoTo AfficherErreur

10      If m_bMonthViewHasFocus = True Then
15        Call cmdTransport_Click
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmCédule", "cmdTransport_MouseUp", Err, Erl
End Sub

Private Sub Form_Resize()

5       On Error GoTo AfficherErreur

10      Call frasemaine.Refresh

15      Call frajour.Refresh

20      Exit Sub

AfficherErreur:

25      woups "frmCédule", "Form_Resize", Err, Erl
End Sub

Private Sub chkfin_Click()

5       On Error GoTo AfficherErreur

        'On change l'affichage sur click
        'Fin projet ou transport
10      If chkfin.Value = vbUnchecked Then
15        Call AfficherTransport
20      Else
25        Call AfficherProjet
30      End If

35      m_bMonthViewHasFocus = False

40      Exit Sub

AfficherErreur:

45      woups "frmCédule", "chkfin_Click", Err, Erl
End Sub

Private Sub AfficherTransport()

5       On Error GoTo AfficherErreur

        'Affichage transport
10      lblprojet.Visible = False
15      txtnoprojet.Visible = False
20      cmbtransport.Visible = True
25      cmdtransport.Visible = True
30      lbltransport.Visible = True

35      Exit Sub

AfficherErreur:

40      woups "frmCédule", "AfficherTransport", Err, Erl
End Sub

Private Sub AfficherProjet()

5       On Error GoTo AfficherErreur
              
        'Affichage fin de projet
10      lblprojet.Visible = True
15      txtnoprojet.Visible = True
20      cmbtransport.Visible = False
25      cmdtransport.Visible = False
30      lbltransport.Visible = False

35      Exit Sub

AfficherErreur:

40      woups "frmCédule", "AfficherProjet", Err, Erl
End Sub

Private Sub cmdAjouterCédule_Click()

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer

        'Met en mode ajouter et affiche champ pour entrer des données
15      m_bModeAjouter = True

20      fraliste.Visible = False
25      fraAlarme.Visible = False
30      frajour.Visible = True

        'Vide champ text
35      cmbemployé.Text = vbNullString

40      For iCompteur = 0 To cmbheuredébut.ListCount - 1
45        If cmbheuredébut.LIST(iCompteur) = "8:00" Then
50          cmbheuredébut.ListIndex = iCompteur

55          Exit For
60        End If
65      Next

70      For iCompteur = 0 To cmbheurefin.ListCount - 1
75        If cmbheurefin.LIST(iCompteur) = "17:00" Then
80          cmbheurefin.ListIndex = iCompteur

85          Exit For
90        End If
95      Next

100     cmbtransport.Text = vbNullString
105     txtnoprojet.Text = vbNullString
110     cmbclient.Text = vbNullString
115     chkfin = vbUnchecked

120     m_bMonthViewHasFocus = False

125     Exit Sub

AfficherErreur:

130     woups "frmCédule", "cmdAjouter_Click", Err, Erl
End Sub

Private Sub cmdAjouterAlarme_Click()

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer

        'Met en mode ajouter et affiche champ pour entrer des données
15      m_bModeAjouter = True

20      mskHeure.Text = ""
25      txtMessage.Text = ""

30      fraliste.Visible = False
35      fraAlarme.Visible = True
40      frajour.Visible = False

45      mskHeure.Text = ""

50      m_bMonthViewHasFocus = False

55      Exit Sub

AfficherErreur:

60      woups "frmCédule", "cmdAjouter_Click", Err, Erl
End Sub

Private Sub cmdAnnulerAlarme_Click()

5       On Error GoTo AfficherErreur

        'Quitte écran pour ajouter ou modifier
10      fraliste.Visible = True
15      fraAlarme.Visible = False
20      frajour.Visible = False

25      m_bMonthViewHasFocus = False

30      Exit Sub

AfficherErreur:

35      woups "frmCédule", "cmdAnnulerAlarme_Click", Err, Erl
End Sub

Private Sub cmdAnnulerCédule_Click()

5       On Error GoTo AfficherErreur

        'quitte ecran pour ajouté ou modifié
10      fraliste.Visible = True
15      fraAlarme.Visible = False
20      frajour.Visible = False

25      m_bMonthViewHasFocus = False

30      Exit Sub

AfficherErreur:

35      woups "frmCédule", "cmdAnnuler_Click", Err, Erl
End Sub

Private Sub CopierAlarme(ByVal datDate As Date)

5       On Error GoTo AfficherErreur

10      Dim sDate          As String
15      Dim rstAlarme      As ADODB.Recordset
20      Dim rstCopieAlarme As ADODB.Recordset
25      Dim iCompteur      As Integer
    
30      sDate = ConvertDate(datDate)
  
35      If sDate <> vbNullString Then
40        datDate = DateSerial(Left$(sDate, 4), Mid$(sDate, 6, 2), Right$(sDate, 2))
    
45        For iCompteur = 1 To Lstjour.ListItems.count
50          If Lstjour.ListItems(iCompteur).Selected = True Then
              'ouvre la table
55            Set rstAlarme = New ADODB.Recordset
60            Set rstCopieAlarme = New ADODB.Recordset

65            Call rstAlarme.Open("SELECT * FROM GRB_Alarmes WHERE IDAlarme = " & Lstjour.ListItems(iCompteur).Text & " ORDER BY Initiale", g_connData, adOpenDynamic, adLockOptimistic)
70            Call rstCopieAlarme.Open("SELECT * FROM GRB_Alarmes", g_connData, adOpenDynamic, adLockOptimistic)
    
75            Call rstCopieAlarme.AddNew
    
              'Ajoute l'enregistrement
80            rstCopieAlarme.Fields("Initiale") = rstAlarme.Fields("Initiale")
85            rstCopieAlarme.Fields("Date") = sDate
90            rstCopieAlarme.Fields("Heure") = rstAlarme.Fields("Heure")
95            rstCopieAlarme.Fields("JourSemaine") = Weekday(datDate)
100           rstCopieAlarme.Fields("Type") = "C"
      
105           Call rstCopieAlarme.Update
                    
              'Quitte l'écran pour ajouté ou modifié
110           fraliste.Visible = True
115           fraAlarme.Visible = False
120           frajour.Visible = False
            
125           Call rstAlarme.Close
130           Set rstAlarme = Nothing
                      
135           Call rstCopieAlarme.Close
140           Set rstCopieAlarme = Nothing
145         End If
150       Next
                      
          'Met à jour l'écran
155       Call RemplirFinProjet
160       Call RemplirListerJour
165       Call RemplirListerSemaine
170     End If

175     Exit Sub

AfficherErreur:

180     woups "frmCédule", "CopierCedule", Err, Erl
End Sub

Private Sub CopierCédule(ByVal datDate As Date)

5       On Error GoTo AfficherErreur

10      Dim sDate          As String
15      Dim rstCédule      As ADODB.Recordset
20      Dim rstCopieCédule As ADODB.Recordset
25      Dim iCompteur      As Integer
    
30      sDate = ConvertDate(datDate)
  
35      If sDate <> vbNullString Then
40        datDate = DateSerial(Left$(sDate, 4), Mid$(sDate, 6, 2), Right$(sDate, 2))
    
45        For iCompteur = 1 To Lstjour.ListItems.count
50         If Lstjour.ListItems(iCompteur).Selected = True Then
              'ouvre la table
55            Set rstCédule = New ADODB.Recordset
60            Set rstCopieCédule = New ADODB.Recordset
              
65            Call rstCédule.Open("SELECT * FROM GRB_cédule WHERE noenreg = " & Lstjour.ListItems(iCompteur).Text & " ORDER BY initiale", g_connData, adOpenDynamic, adLockOptimistic)
70            Call rstCopieCédule.Open("SELECT * FROM GRB_cédule", g_connData, adOpenDynamic, adLockOptimistic)
    
75            Call rstCopieCédule.AddNew
    
              ''''''''''''''''''''''''''
              'ajoute l'enregistrement
              ''''''''''''''''''''''''''
80            rstCopieCédule.Fields("initiale") = rstCédule.Fields("initiale")
85            rstCopieCédule.Fields("date_cedulé") = sDate
90            rstCopieCédule.Fields("heure_début") = rstCédule.Fields("heure_début")
95            rstCopieCédule.Fields("heure_fin") = rstCédule.Fields("heure_fin")
100           rstCopieCédule.Fields("Client") = rstCédule.Fields("Client")
105           rstCopieCédule.Fields("joursemaine") = Weekday(datDate)
110           rstCopieCédule.Fields("finprojet") = rstCédule.Fields("finprojet")
115           rstCopieCédule.Fields("transport") = rstCédule.Fields("transport")
        
120           Call rstCopieCédule.Update
                    
              'quitte l'écran pour ajouté ou modifié
125           fraliste.Visible = True
130           frajour.Visible = False
            
135           Call rstCédule.Close
140           Set rstCédule = Nothing

145           Call rstCopieCédule.Close
150           Set rstCopieCédule = Nothing
155         End If
160       Next
                      
          'met a jour l'écran
165       Call RemplirFinProjet
170       Call RemplirListerJour
175       Call RemplirListerSemaine
180     End If

185     Exit Sub

AfficherErreur:

190     woups "frmCédule", "CopierCedule", Err, Erl
End Sub

Private Sub cmdCopier_Click()

5       On Error GoTo AfficherErreur

10      If Lstjour.ListItems.count > 0 Then
15        mvwChoixDate.Month = mvwJanuary
20        mvwChoixDate.Day = 1
  
25        mvwChoixDate.Year = mvwSelection.Year
30        mvwChoixDate.Month = mvwSelection.Month
35        mvwChoixDate.Day = mvwSelection.Day
  
40        mvwChoixDate.Visible = True
  
45        Call mvwChoixDate.SetFocus
50      End If

55      m_bMonthViewHasFocus = False

60      Exit Sub

AfficherErreur:

65      woups "frmCédule", "cmdCopier_Click", Err, Erl
End Sub

Private Sub cmdEnregistrerAlarme_Click()

5       On Error GoTo AfficherErreur

        'Enregistre
10      Dim rstAlarme  As ADODB.Recordset
15      Dim rstEmploye As ADODB.Recordset
20      Dim iNoEmploye As Integer

25      If mskHeure.Text <> "" Then
30        If IsDate(mskHeure.Text) Then
35          If txtMessage.Text <> "" Then
40            Set rstAlarme = New ADODB.Recordset
45            Set rstEmploye = New ADODB.Recordset

              'Ouvre la table
50            If m_bModeAjouter = True Then
55              Call rstAlarme.Open("SELECT * FROM GRB_Alarmes", g_connData, adOpenDynamic, adLockOptimistic)
  
60              m_bModeAjouter = False

65              Call rstAlarme.AddNew
70            Else
75              Call rstAlarme.Open("SELECT * FROM GRB_Alarmes WHERE IDAlarme = " & Lstjour.SelectedItem.Text, g_connData, adOpenDynamic, adLockOptimistic)
80            End If

85            Call rstEmploye.Open("SELECT * FROM GRB_Employés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)

90            iNoEmploye = rstEmploye.Fields("NoEmploye")
  
95            Call rstEmploye.Close
100           Set rstEmploye = Nothing
  
              'Ajoute l'enregistrement
105           rstAlarme.Fields("NoEmploye") = iNoEmploye
110           rstAlarme.Fields("Date") = ConvertDate(m_datDateChoisie)
115           rstAlarme.Fields("Heure") = mskHeure.Text
120           rstAlarme.Fields("Message") = txtMessage.Text
125           rstAlarme.Fields("JourSemaine") = Weekday(m_datDateChoisie)
130           rstAlarme.Fields("TypeCédule") = "C"
       
              'Quitte ecran pour ajouté ou modifié
135           fraliste.Visible = True
140           fraAlarme.Visible = False
145           frajour.Visible = False
            
150           Call rstAlarme.Update
            
155           Call rstAlarme.Close
160           Set rstAlarme = Nothing
                      
              'Met à jour l'écran
165           Call RemplirFinProjet
170           Call RemplirListerJour
175           Call RemplirListerSemaine
180         Else
185           Call MsgBox("Il n'y a pas de message à afficher!", vbOKOnly, "Erreur")
190         End If
195       Else
200         Call MsgBox("L'heure est invalide!", vbOKOnly, "Erreur")
205       End If
210     Else
215       Call MsgBox("L'heure est obligatoire!", vbOKOnly, "Erreur")
220     End If

225     m_bMonthViewHasFocus = False

230     Exit Sub

AfficherErreur:

235     woups "frmCédule", "cmdEnregistrerAlarme_Click", Err, Erl
End Sub

Private Sub cmdEnregistrerCédule_Click()

5       On Error GoTo AfficherErreur

        'Enregistre
10      Dim rstCédule  As ADODB.Recordset
15      Dim rstEmployé As ADODB.Recordset
  
20      If cmbemployé.ListIndex <> -1 Then
25        Set rstCédule = New ADODB.Recordset
30        Set rstEmployé = New ADODB.Recordset

          'Ouvre la table
35        If m_bModeAjouter = True Then
40          Call rstCédule.Open("SELECT * FROM GRB_cédule", g_connData, adOpenDynamic, adLockOptimistic)
      
45          m_bModeAjouter = False
      
50          Call rstCédule.AddNew
55        Else
60          Call rstCédule.Open("SELECT * FROM GRB_cédule WHERE noenreg = " & Lstjour.ListItems(Lstjour.SelectedItem.Index).Text & " ORDER BY initiale", g_connData, adOpenDynamic, adLockOptimistic)
65        End If
           
          'Ajoute l'enregistrement
70        Call rstEmployé.Open("SELECT initiale FROM GRB_employés WHERE noemploye = " & cmbemployé.ItemData(cmbemployé.ListIndex), g_connData, adOpenDynamic, adLockOptimistic)
        
75        rstCédule.Fields("initiale") = rstEmployé.Fields("initiale")
        
80        Call rstEmployé.Close
85        Set rstEmployé = Nothing
            
90        rstCédule.Fields("date_cedulé") = ConvertDate(m_datDateChoisie)
            
95        If cmbheuredébut.Text = vbNullString Then
100         rstCédule.Fields("heure_début") = " "
110       Else
115         rstCédule.Fields("heure_début") = cmbheuredébut.Text
120       End If
            
125       If cmbheurefin.Text = vbNullString Then
130         rstCédule.Fields("heure_fin") = " "
135       Else
140         rstCédule.Fields("heure_fin") = cmbheurefin.Text
145       End If
           
150       If cmbclient.Text = vbNullString Then
155         rstCédule.Fields("CLIENT") = " "
160       Else
165         rstCédule.Fields("CLIENT") = cmbclient.Text
170       End If
            
175       rstCédule.Fields("joursemaine") = Weekday(m_datDateChoisie)
            
180       rstCédule.Fields("finprojet") = chkfin.Value
       
          'Enregistre le champ finprojet ou transport
185       If chkfin.Value = vbUnchecked Then
190         If cmbtransport.Text = vbNullString Then
195           rstCédule.Fields("transport") = " "
200         Else
205           rstCédule.Fields("transport") = cmbtransport.Text
210         End If
215       Else
220         If txtnoprojet.Text = vbNullString Then
225           rstCédule.Fields("transport") = " "
230         Else
235           rstCédule.Fields("transport") = txtnoprojet.Text
240         End If
245       End If
     
250       Call rstCédule.Update
                   
          'Quitte ecran pour ajouté ou modifié
255       fraliste.Visible = True
260       fraAlarme.Visible = False
265       frajour.Visible = False
            
270       Call rstCédule.Close
275       Set rstCédule = Nothing
                      
          'Met à jour l'écran
280       Call RemplirFinProjet
285       Call RemplirListerJour
290       Call RemplirListerSemaine
295     Else
300       Call MsgBox("Aucun employé de sélectionné!")
305     End If

310     m_bMonthViewHasFocus = False

315     Exit Sub

AfficherErreur:

320     woups "frmCédule", "cmdEnregistrerCédule_Click", Err, Erl
End Sub

Private Sub cmdsupprimer_Click()

5       On Error GoTo AfficherErreur
        
        'Supprime la cédule selectionnée
10      Dim iCompteur As Integer
     
15      If Lstjour.ListItems.count > 0 Then
20        If MsgBox("Voulez-vous supprimer ce(ces) rendez-vous?", vbYesNo) = vbYes Then
25          For iCompteur = 1 To Lstjour.ListItems.count
30            If Lstjour.ListItems(iCompteur).Selected = True Then
35              If Lstjour.ListItems(iCompteur).Tag = "A" Then
40                Call g_connData.Execute("DELETE * FROM GRB_Alarmes WHERE IDAlarme = " & Lstjour.ListItems(iCompteur).Text)
45              Else
50                Call g_connData.Execute("DELETE * FROM GRB_Cédule WHERE noenreg = " & Lstjour.ListItems(iCompteur).Text)
55              End If
60            End If
65          Next
        
            'Mise à jour des ListViews
70          Call RemplirFinProjet
75          Call RemplirListerJour
80          Call RemplirListerSemaine
85        End If
90      End If

95      m_bMonthViewHasFocus = False

100     Exit Sub

AfficherErreur:

105     woups "frmCédule", "cmdSupprimer_Click", Err, Erl
End Sub

Private Sub cmdTransport_Click()

5       On Error GoTo AfficherErreur

10      Dim rstTransport As ADODB.Recordset
15      Dim sTransport   As String
20      Dim iCompteur    As Integer
    
25      sTransport = cmbtransport.Text
        
        'Si'l y a un transport
30      If cmbtransport.Text <> vbNullString Then
          'Si le transport existe, on demande si on veut le supprimer
          'sinon, on demande si on veut l'ajouter
35        If ComboContient(cmbtransport, sTransport) Then
            'Si réponse oui pour supprimer
40          If MsgBox("Voulez-vous supprimer le transport " & cmbtransport.Text & "?", vbYesNo) = vbYes Then
45            Call g_connData.Execute("DELETE * FROM GRB_transport WHERE transport = '" & Replace(cmbtransport.Text, "'", "''") & "'")
50          Else
              'Sinon demande si veut ajouter un nouveau transport
55            If MsgBox("Voulez-vous ajouter un transport?", vbYesNo) = vbYes Then
60              sTransport = InputBox("Veuillez entrer son nom!")
          
                'Si quelque chose d'entrer
65              If sTransport <> vbNullString Then
70                If Not ComboContient(cmbtransport, sTransport) Then
75                  Set rstTransport = New ADODB.Recordset

80                  Call rstTransport.Open("SELECT * FROM GRB_Transport", g_connData, adOpenDynamic, adLockOptimistic)
              
85                  Call rstTransport.AddNew

90                  rstTransport.Fields("transport").Value = sTransport
                
95                  Call rstTransport.Update
             
100                 Call rstTransport.Close
105                 Set rstTransport = Nothing
110               Else
115                 Call MsgBox("Ce transport existe déjà!")
120               End If
125             End If
130           End If
135         End If
140       Else
            'Demande si veut ajouter un nouveau transport
145         If MsgBox("Voulez-vous ajouter un transport?", vbYesNo) = vbYes Then
150           sTransport = InputBox("Veuillez entrer son nom!")
        
              'Si quelque chose d'entrer
155           If sTransport <> vbNullString Then
160             If Not ComboContient(cmbtransport, sTransport) Then
165               Set rstTransport = New ADODB.Recordset

170               Call rstTransport.Open("SELECT * FROM GRB_transport", g_connData, adOpenDynamic, adLockOptimistic)

175               Call rstTransport.AddNew
180               rstTransport.Fields("transport") = sTransport
              
185               Call rstTransport.Update
          
190               Call rstTransport.Close
195               Set rstTransport = Nothing
200             Else
205               Call MsgBox("Ce transport existe déjà!")
210             End If
215           End If
220         End If
225       End If
230     Else
          'Demande si veut ajouter un nouveau transport
235       If MsgBox("Voulez-vous ajouter un transport?", vbYesNo) = vbYes Then
240         sTransport = InputBox("Veuillez entrer son nom!")
      
            'Si quelque chose d'entrer
245         If sTransport <> vbNullString Then
250           If Not ComboContient(cmbtransport, sTransport) Then
255             Set rstTransport = New ADODB.Recordset

260             Call rstTransport.Open("SELECT * FROM GRB_transport", g_connData, adOpenDynamic, adLockOptimistic)
        
265             Call rstTransport.AddNew

270             rstTransport.Fields("transport") = sTransport
          
275             Call rstTransport.Update
          
280             Call rstTransport.Close
285             Set rstTransport = Nothing
290           Else
295             Call MsgBox("Ce transport existe déjà!")
300           End If
305         End If
310       End If
315     End If
    
        'Remplis combo transport
320     Call RemplirTransport
  
325     For iCompteur = 0 To cmbtransport.ListCount - 1
330       If cmbtransport.LIST(iCompteur) = sTransport Then
335         cmbtransport.ListIndex = iCompteur
      
340         Exit For
345       End If
350     Next
  
355     If cmbtransport.ListIndex = -1 Then
360       cmbtransport.ListIndex = 0
365     End If

370     m_bMonthViewHasFocus = False

375     Exit Sub

AfficherErreur:

380     woups "frmCédule", "cmdtransport_Click", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer

15      mvwSelection.StartOfWeek = mvwSunday

20      g_bCeduleOuverte = True

        'Met à jour l'écran
25      mvwSelection.Year = Year(Date)
30      mvwSelection.Month = Month(Date)
35      mvwSelection.Day = Day(Date)

40      m_datDateChoisie = Date
  
        'Rempli les combos
45      Call RemplirListerEmployé
50      Call RemplirTransport
55      Call RemplirListerClient
60      Call RemplirFinProjet
  
        'Rempli les ListViews
65      Call RemplirListerJour
70      Call RemplirListerSemaine
  
        'Sélectionne le jour de la semaine
75      For iCompteur = 1 To 7
80        If lstjoursemaine(iCompteur).Tag = m_datDateChoisie Then
85          lstjoursemaine(iCompteur).BackColor = &HE0E0E0
90        Else
95          lstjoursemaine(iCompteur).BackColor = &HFFFFFF
100       End If
105     Next
    
110     Call AfficherTransport

115     Screen.MousePointer = vbDefault

120     Exit Sub

AfficherErreur:

125     woups "frmCédule", "Form_Load", Err, Erl
End Sub

Private Sub RemplirFinProjet()

5       On Error GoTo AfficherErreur
        
        'Remplis le combo transport
10      Dim rstCedule As ADODB.Recordset
  
15      Set rstCedule = New ADODB.Recordset
  
20      Call rstCedule.Open("SELECT date_cedulé, transport FROM GRB_cédule WHERE finprojet = 1 ORDER BY date_cedulé", g_connData, adOpenDynamic, adLockOptimistic)
    
25      Call cmbfinprojet.Clear
    
        'Rempli tant il y a des employé
30      Do While Not rstCedule.EOF
35        DoEvents
    
40        Call cmbfinprojet.AddItem(Trim$(CStr(rstCedule!transport)) & "     " & ConvertDate(rstCedule!date_cedulé))
      
45        Call rstCedule.MoveNext
50      Loop
    
55      Call rstCedule.Close
60      Set rstCedule = Nothing

        'S'il y a des enregistrements, on sélectionne le premier
65      If cmbfinprojet.ListCount > 0 Then
70        cmbfinprojet.ListIndex = 0
75      End If

80      Exit Sub

AfficherErreur:

85      woups "frmCédule", "RemplirFinProjet", Err, Erl
End Sub

Private Sub RemplirTransport()

5       On Error GoTo AfficherErreur
        
        ''''''''''''''''''''''''
        'remplis combo transport
        ''''''''''''''''''''''''
10      Dim rstTransport As ADODB.Recordset
  
15      Set rstTransport = New ADODB.Recordset
  
20      Call rstTransport.Open("SELECT * FROM GRB_transport ORDER BY transport", g_connData, adOpenDynamic, adLockOptimistic)
    
25      Call cmbtransport.Clear
    
        'rempli tant il y a des employé
30      Do While Not rstTransport.EOF
35        Call cmbtransport.AddItem(rstTransport!transport)
     
40        Call rstTransport.MoveNext
45      Loop
    
50      Call rstTransport.Close
55      Set rstTransport = Nothing

60      Exit Sub

AfficherErreur:

65      woups "frmCédule", "RemplirTransport", Err, Erl
End Sub

Private Sub RemplirListerEmployé()

5       On Error GoTo AfficherErreur
        
        '''''''''''''''''''''''''
        ' Remplis combo employé '
        '''''''''''''''''''''''''
10      Dim rstEmploye As ADODB.Recordset

15      Set rstEmploye = New ADODB.Recordset

20      Call rstEmploye.Open("SELECT * FROM GRB_employés WHERE Actif = True ORDER BY employe", g_connData, adOpenDynamic, adLockOptimistic)
    
25      Call cmbemployé.Clear
    
              'rempli tant il y a des employé
30      Do While Not rstEmploye.EOF
35        Call cmbemployé.AddItem(rstEmploye!employe)
40        cmbemployé.ItemData(cmbemployé.newIndex) = rstEmploye!noEmploye
        
45        Call rstEmploye.MoveNext
50      Loop
    
55      Call rstEmploye.Close
60      Set rstEmploye = Nothing

65      Exit Sub

AfficherErreur:

70      woups "frmCédule", "RemplirListerEmployé", Err, Erl
End Sub

Private Sub RemplirListerClient()

5       On Error GoTo AfficherErreur
        
        'remplis combo employé
10      Dim rstClient As ADODB.Recordset

15      Set rstClient = New ADODB.Recordset

20      Call rstClient.Open("SELECT * FROM GRB_Client WHERE Supprimé = False ORDER BY NomClient", g_connData, adOpenDynamic, adLockOptimistic)
    
25      Call cmbclient.Clear
    
        'rempli tant il y a des employé
30      Do While Not rstClient.EOF
35        Call cmbclient.AddItem(rstClient!nomclient)
        
40        Call rstClient.MoveNext
45      Loop
        
50      Call rstClient.Close
55      Set rstClient = Nothing

60      Exit Sub

AfficherErreur:

70      woups "frmCédule", "RemplirListerClient", Err, Erl
End Sub

Public Sub RemplirListerJour()

5       On Error GoTo AfficherErreur

        'Remplis lister une journée
10      Dim rstCédule As ADODB.Recordset
15      Dim itmCedule As ListItem
  
        'Vide le lister
20      Call Lstjour.ListItems.Clear
  
25      Set rstCédule = New ADODB.Recordset
  
30      Call rstCédule.Open("SELECT * FROM GRB_cédule WHERE date_cedulé = '" & ConvertDate(m_datDateChoisie) & "' ORDER BY initiale", g_connData, adOpenDynamic, adLockOptimistic)
           
        'Tant il y a de employé cedulé , ajoute dans lister
35      Do While Not rstCédule.EOF
40        Set itmCedule = Lstjour.ListItems.Add
            
45        itmCedule.Text = rstCédule!noenreg

50        itmCedule.Tag = "C"

55        If Not IsNull(rstCédule.Fields("Initiale")) Then
60          itmCedule.SubItems(I_LVW_JOUR_NOM) = rstCédule!Initiale
65        Else
70          itmCedule.SubItems(I_LVW_JOUR_NOM) = ""
75        End If
            
80        If Not IsNull(rstCédule!heure_début) Then
85          itmCedule.SubItems(I_LVW_JOUR_DEBUT) = rstCédule!heure_début
90        Else
95          itmCedule.SubItems(I_LVW_JOUR_DEBUT) = ""
100       End If
            
105       If Not IsNull(rstCédule!heure_fin) Then
110         itmCedule.SubItems(I_LVW_JOUR_FIN) = rstCédule!heure_fin
115       Else
120         itmCedule.SubItems(I_LVW_JOUR_FIN) = ""
125       End If
            
130       If Not IsNull(rstCédule!CLIENT) Then
135         itmCedule.SubItems(I_LVW_JOUR_CLIENT) = rstCédule!CLIENT
140       Else
145         itmCedule.SubItems(I_LVW_JOUR_CLIENT) = ""
150       End If
            
          'si fin de projet marque numero de projet sinon transport
155       If rstCédule!finprojet = 0 Then
160         If Not IsNull(rstCédule!transport) Then
165           itmCedule.SubItems(I_LVW_JOUR_TRANSPORT) = rstCédule!transport
170         Else
175           itmCedule.SubItems(I_LVW_JOUR_TRANSPORT) = ""
180         End If
              
            'met en rouge ou en noir dépendant si fin de projet
185         itmCedule.ListSubItems(I_LVW_JOUR_NOM).ForeColor = COLOR_NOIR
190         itmCedule.ListSubItems(I_LVW_JOUR_DEBUT).ForeColor = COLOR_NOIR
195         itmCedule.ListSubItems(I_LVW_JOUR_FIN).ForeColor = COLOR_NOIR
200         itmCedule.ListSubItems(I_LVW_JOUR_CLIENT).ForeColor = COLOR_NOIR
205         itmCedule.ListSubItems(I_LVW_JOUR_TRANSPORT).ForeColor = COLOR_NOIR
210       Else
215         If Not IsNull(rstCédule!transport) Then
220           itmCedule.SubItems(I_LVW_JOUR_TRANSPORT) = "Fin " + rstCédule!transport
225         Else
230           itmCedule.SubItems(I_LVW_JOUR_TRANSPORT) = "Fin"
235         End If
               
            'Met en rouge ou en noir dépendant si fin de projet
240         itmCedule.ListSubItems(I_LVW_JOUR_NOM).ForeColor = COLOR_ROUGE
245         itmCedule.ListSubItems(I_LVW_JOUR_DEBUT).ForeColor = COLOR_ROUGE
250         itmCedule.ListSubItems(I_LVW_JOUR_FIN).ForeColor = COLOR_ROUGE
255         itmCedule.ListSubItems(I_LVW_JOUR_CLIENT).ForeColor = COLOR_ROUGE
260         itmCedule.ListSubItems(I_LVW_JOUR_TRANSPORT).ForeColor = COLOR_ROUGE
265       End If
            
270       Call rstCédule.MoveNext
275     Loop
    
280     Call rstCédule.Close
285     Set rstCédule = Nothing

290     Call RemplirListerJourAlarme

295     Exit Sub

AfficherErreur:

300     woups "frmCédule", "RemplirListerJour", Err, Erl
End Sub

Public Sub RemplirListerSemaine()

5       On Error GoTo AfficherErreur
              
        'Remplis une semaine
10      Dim rstCédule       As ADODB.Recordset
15      Dim iJourSemaine    As Integer
20      Dim datPremiereDate As Date
25      Dim datDerniereDate As Date
30      Dim iCompteur       As Integer
35      Dim sHeureDebutFin  As String
40      Dim itmSemaine      As ListItem

45      For iCompteur = 1 To 7
          'couleur par defaut entete de date
50        lbljour(iCompteur - 1).ForeColor = vbWhite
55        lbljourstr(iCompteur - 1).ForeColor = vbWhite

60        Call lstjoursemaine(iCompteur).ListItems.Clear
65      Next
    
70      iJourSemaine = Weekday(m_datDateChoisie)
75      datPremiereDate = m_datDateChoisie
80      datDerniereDate = m_datDateChoisie
    
        'trouve premiere date de la semaine
85      Do While Not Weekday(datPremiereDate) = 1
90        datPremiereDate = datPremiereDate - 1
95      Loop
    
        'trouve derniere date de la semaine
100     Do While Not Weekday(datDerniereDate) = 7
105       datDerniereDate = datDerniereDate + 1
110     Loop
    
        'selectionne la semaine courante
115     Set rstCédule = New ADODB.Recordset
        
120     Call rstCédule.Open("SELECT * FROM GRB_cédule WHERE cdate(date_cedulé) <= cdate('" & CStr(datDerniereDate) & "') AND cdate(date_cedulé) >= cdate('" & CStr(datPremiereDate) & "') ORDER BY initiale", g_connData, adOpenDynamic, adLockOptimistic)
         
125     For iCompteur = 1 To 7
          'pour ecrire le jour
130       lbljour(iCompteur - 1).Caption = Day(datPremiereDate + iCompteur - 1)
      
          'garde en memoire la date des lister
135       lstjoursemaine(iCompteur).Tag = datPremiereDate + iCompteur - 1
140     Next
    
145     Do While Not rstCédule.EOF
          'ajoute dans le lister, dépendant le jour de la semaine
150       Set itmSemaine = lstjoursemaine(rstCédule!joursemaine).ListItems.Add
     
155       itmSemaine.Text = rstCédule!noenreg
       
          'si fin de projet marque numero de projet sinon transport
160       If rstCédule!finprojet = 0 Then
165         If Not IsNull(rstCédule.Fields("Initiale")) Then
170           itmSemaine.SubItems(I_LVW_SEMAINE_NOM) = rstCédule!Initiale
175         Else
180           itmSemaine.SubItems(I_LVW_SEMAINE_NOM) = ""
185         End If
               
190         If Not IsNull(rstCédule!heure_début) Then
195           sHeureDebutFin = Trim(rstCédule!heure_début + "-")
200         Else
205           sHeureDebutFin = "-"
210         End If
          
215         If Not IsNull(rstCédule!heure_fin) Then
220           sHeureDebutFin = sHeureDebutFin + rstCédule!heure_fin
225         Else
230           sHeureDebutFin = sHeureDebutFin + " "
235         End If
          
240         itmSemaine.SubItems(I_LVW_SEMAINE_HEURE) = sHeureDebutFin
        
            'met en noir
245         itmSemaine.ListSubItems(I_LVW_SEMAINE_NOM).ForeColor = COLOR_NOIR
250         itmSemaine.ListSubItems(I_LVW_SEMAINE_HEURE).ForeColor = COLOR_NOIR
255       Else
260         itmSemaine.SubItems(I_LVW_SEMAINE_NOM) = "Fin"
              
265         lbljour(rstCédule!joursemaine - 1).ForeColor = COLOR_ROUGE
270         lbljourstr(rstCédule!joursemaine - 1).ForeColor = COLOR_ROUGE
              
275         If Not IsNull(rstCédule!transport) Then
280           sHeureDebutFin = rstCédule!transport
285         Else
290           sHeureDebutFin = " "
295         End If
              
300         itmSemaine.SubItems(I_LVW_SEMAINE_HEURE) = sHeureDebutFin
        
            'met en rouge
305         itmSemaine.ListSubItems(I_LVW_SEMAINE_NOM).ForeColor = COLOR_ROUGE
310         itmSemaine.ListSubItems(I_LVW_SEMAINE_HEURE).ForeColor = COLOR_ROUGE
315       End If
        
320       Call rstCédule.MoveNext
325     Loop
        
330     Call rstCédule.Close
335     Set rstCédule = Nothing

340     Call RemplirListerSemaineAlarme

345     Exit Sub

AfficherErreur:

350     woups "frmCédule", "RemplirListerSemaine", Err, Erl
End Sub

Private Sub RemplirListerJourAlarme()

5       On Error GoTo AfficherErreur

        'Remplis lister une journée
10      Dim rstAlarme  As ADODB.Recordset
15      Dim rstEmploye As ADODB.Recordset
20      Dim iNoEmploye As Integer
25      Dim itmAlarme  As ListItem

30      Set rstEmploye = New ADODB.Recordset

35      Call rstEmploye.Open("SELECT * FROM GRB_Employés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)
    
40      iNoEmploye = rstEmploye.Fields("NoEmploye")

45      Call rstEmploye.Close
50      Set rstEmploye = Nothing

55      Set rstAlarme = New ADODB.Recordset

60      Call rstAlarme.Open("SELECT * FROM GRB_Alarmes WHERE Date = '" & ConvertDate(m_datDateChoisie) & "' AND NoEmploye = " & iNoEmploye & " AND TypeCédule = 'C' ORDER BY Date, Heure", g_connData, adOpenDynamic, adLockOptimistic)
           
        'Tant il y a de employé cedulé , ajoute dans lister
65      Do While Not rstAlarme.EOF
70        Set itmAlarme = Lstjour.ListItems.Add
            
75        itmAlarme.Text = (rstAlarme.Fields("IDAlarme"))
80        itmAlarme.Tag = "A"

85        itmAlarme.SubItems(I_LVW_JOUR_NOM) = g_sInitiale
            
90        If Not IsNull(rstAlarme.Fields("Heure")) Then
95          itmAlarme.SubItems(I_LVW_JOUR_DEBUT) = rstAlarme.Fields("Heure")
100       Else
105         itmAlarme.SubItems(I_LVW_JOUR_DEBUT) = ""
110       End If
                       
          'Met en rouge ou en noir dépendant si fin de projet
115       itmAlarme.ListSubItems(I_LVW_JOUR_NOM).ForeColor = COLOR_BLEU
120       itmAlarme.ListSubItems(I_LVW_JOUR_DEBUT).ForeColor = COLOR_BLEU
          
125       Call rstAlarme.MoveNext
130     Loop
    
135     Call rstAlarme.Close
140     Set rstAlarme = Nothing

145     Exit Sub

AfficherErreur:

150     woups "frmCédule", "RemplirListerJourAlarme", Err, Erl
End Sub

Private Sub RemplirListerSemaineAlarme()

5       On Error GoTo AfficherErreur

        'Remplis une semaine
10      Dim rstAlarme       As ADODB.Recordset
15      Dim rstEmploye      As ADODB.Recordset
20      Dim iNoEmploye      As Integer
25      Dim iJourSemaine    As Integer
30      Dim datPremiereDate As Date
35      Dim datDerniereDate As Date
40      Dim iCompteur       As Integer
45      Dim itmSemaine      As ListItem
    
50      iJourSemaine = Weekday(m_datDateChoisie)
55      datPremiereDate = m_datDateChoisie
60      datDerniereDate = m_datDateChoisie
    
        'Trouve première date de la semaine
65      Do While Not Weekday(datPremiereDate) = 1
70        datPremiereDate = datPremiereDate - 1
75      Loop
    
        'Trouve dernière date de la semaine
80      Do While Not Weekday(datDerniereDate) = 7
85        datDerniereDate = datDerniereDate + 1
90      Loop

95      Set rstEmploye = New ADODB.Recordset

100     Call rstEmploye.Open("SELECT * FROM GRB_Employés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)

105     iNoEmploye = rstEmploye.Fields("NoEmploye")

110     Call rstEmploye.Close
115     Set rstEmploye = Nothing
    
        'Sélectionne la semaine courante
120     Set rstAlarme = New ADODB.Recordset
        
125     Call rstAlarme.Open("SELECT * FROM GRB_Alarmes WHERE Date BETWEEN '" & ConvertDate(datPremiereDate) & "' AND '" & ConvertDate(datDerniereDate) & "' AND NoEmploye = " & iNoEmploye & " AND TypeCédule = 'C' ORDER BY Date, Heure", g_connData, adOpenDynamic, adLockOptimistic)
    
        'iSemaine est le numero du lister, jour de semaine
130     For iCompteur = 1 To 7
          'pour ecrire le jour
135       lbljour(iCompteur - 1).Caption = Day(datPremiereDate + iCompteur - 1)
      
          'garde en memoire la date des lister
140       lstjoursemaine(iCompteur).Tag = datPremiereDate + iCompteur - 1
145     Next
    
150     Do While Not rstAlarme.EOF
          'ajoute dans le lister, dépendant le jour de la semaine
155       Set itmSemaine = lstjoursemaine(rstAlarme.Fields("JourSemaine")).ListItems.Add
      
160       itmSemaine.Text = rstAlarme.Fields("IDAlarme")
       
          'si fin de projet marque numero de projet sinon transport
165       itmSemaine.SubItems(I_LVW_SEMAINE_NOM) = g_sInitiale
              
170       itmSemaine.SubItems(I_LVW_SEMAINE_HEURE) = rstAlarme.Fields("Heure")
        
          'met en noir
175       itmSemaine.ListSubItems(I_LVW_SEMAINE_NOM).ForeColor = COLOR_BLEU
180       itmSemaine.ListSubItems(I_LVW_SEMAINE_HEURE).ForeColor = COLOR_BLEU
        
185       Call rstAlarme.MoveNext
190     Loop
        
195     Call rstAlarme.Close
200     Set rstAlarme = Nothing

205     Exit Sub

AfficherErreur:

210     woups "frmCédule", "RemplirListerSemaineAlarme", Err, Erl
End Sub

Private Sub Form_Unload(Cancel As Integer)

5       On Error GoTo AfficherErreur

10      g_bCeduleOuverte = False

15      Exit Sub

AfficherErreur:

20      woups "frmCédule", "Form_Unload", Err, Erl
End Sub

Private Sub Lstjour_DblClick()

5       On Error GoTo AfficherErreur

10      Dim rstCédule  As ADODB.Recordset
15      Dim rstAlarme  As ADODB.Recordset
20      Dim rstEmployé As ADODB.Recordset
25      Dim iCompteur  As Integer

        'Affiche en mode modification
30      m_bModeAjouter = False
         
35      If Lstjour.ListItems.count > 0 Then
40        fraliste.Visible = False

45        If Lstjour.SelectedItem.Tag = "C" Then
50          frajour.Visible = True
55          fraAlarme.Visible = False

            'Ouvre la table
60          Set rstCédule = New ADODB.Recordset
65          Set rstEmployé = New ADODB.Recordset
            
70          Call rstCédule.Open("SELECT * FROM GRB_cédule WHERE noenreg = " & Lstjour.SelectedItem.Text & " ORDER BY initiale", g_connData, adOpenDynamic, adLockOptimistic)
75          Call rstEmployé.Open("SELECT * FROM GRB_employés WHERE initiale = '" & rstCédule.Fields("initiale") & "'", g_connData, adOpenDynamic, adLockOptimistic)
    
            'si il y a employé, affiche a l'écran pour modification
80          If Not rstEmployé.EOF Then
85            For iCompteur = 0 To cmbemployé.ListCount - 1
90              If cmbemployé.LIST(iCompteur) = rstEmployé.Fields("Employe") Then
95                cmbemployé.ListIndex = iCompteur

100               Exit For
105             End If
110           Next

115           cmbheuredébut.Text = rstCédule.Fields("heure_début")
120           cmbheurefin.Text = rstCédule.Fields("heure_fin")
        
125           If IsNull(rstCédule!CLIENT) Then
130             cmbclient.Text = " "
135           Else
140             cmbclient.Text = rstCédule!CLIENT
145           End If
        
150           chkfin.Value = rstCédule!finprojet
        
155           If IsNull(rstCédule!transport) Then
160             cmbtransport.Text = " "
165           Else
170             cmbtransport.Text = rstCédule.Fields("transport")
175           End If
        
180           If IsNull(rstCédule!transport) Then
185             txtnoprojet.Text = " "
190           Else
195             txtnoprojet.Text = rstCédule!transport
200           End If
        
              'Affiche fin de projet ou transport
205           If chkfin = vbUnchecked Then
210             Call AfficherTransport
215           Else
220             Call AfficherProjet
225           End If
230         End If
    
235         Call rstCédule.Close
240         Set rstCédule = Nothing
      
245         Call rstEmployé.Close
250         Set rstEmployé = Nothing
255       Else
260         frajour.Visible = False
265         fraAlarme.Visible = True

270         mskHeure.Text = ""
275         txtMessage.Text = ""

            'Ouvre la table
280         Set rstAlarme = New ADODB.Recordset

285         Call rstAlarme.Open("SELECT * FROM GRB_Alarmes WHERE IDAlarme = " & Lstjour.SelectedItem.Text & " ORDER BY NoEmploye", g_connData, adOpenDynamic, adLockOptimistic)
    
290         mskHeure.Text = rstAlarme.Fields("Heure")
       
295         txtMessage.Text = rstAlarme.Fields("Message")
    
300         Call rstAlarme.Close
305         Set rstAlarme = Nothing
310       End If
315     Else
320       fraliste.Visible = True
325       frajour.Visible = False
330     End If

335     Exit Sub

AfficherErreur:

340     woups "frmCédule", "Lstjour_DblClick", Err, Erl
End Sub

Private Sub Lstjour_KeyDown(KeyCode As Integer, Shift As Integer)

5       On Error GoTo AfficherErreur

10      If Lstjour.ListItems.count > 0 Then
15        If KeyCode = vbKeyDelete Then
20          If MsgBox("Voulez-vous supprimer cette enregistrement?", vbYesNo) = vbYes Then
25            If Lstjour.SelectedItem.Tag = "A" Then
30              Call g_connData.Execute("DELETE * FROM GRB_Alarmes WHERE IDAlarme = " & Lstjour.SelectedItem.Text)
35            Else
40              Call g_connData.Execute("DELETE * FROM GRB_Cédule WHERE noenreg = " & Lstjour.SelectedItem.Text)
45            End If

              'Mise à jour des lister
50            Call RemplirFinProjet
55            Call RemplirListerJour
60            Call RemplirListerSemaine
65          End If
70        End If
75      End If

80      Exit Sub

AfficherErreur:

85      woups "frmCédule", "Lstjour_KeyDown", Err, Erl
End Sub

Private Sub lstjoursemaine_Click(Index As Integer)

5       On Error GoTo AfficherErreur

10      Dim iCompteur  As Integer
15      Dim sDate      As String
20      Dim iNbreJour  As Integer
    
        'Initialise la couleur en blanc
25      For iCompteur = 1 To 7
30        lstjoursemaine(iCompteur).BackColor = &HFFFFFF
35      Next
  
        'Sélectionne jour de semaine
40      lstjoursemaine(Index).BackColor = &HE0E0E0

45      sDate = lstjoursemaine(Index).Tag

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
175     Call RemplirListerJour

180     fraliste.Visible = True
185     frajour.Visible = False

190     Call Lstjour.SetFocus

195     Exit Sub

AfficherErreur:

200     woups "frmCédule", "lstjoursemaine_Click", Err, Erl
End Sub

Private Sub mvwChoixDate_DateClick(ByVal DateClicked As Date)

5       On Error GoTo AfficherErreur

10      If Lstjour.SelectedItem.Tag = "A" Then
15        Call CopierAlarme(DateClicked)
20      Else
25        Call CopierCédule(DateClicked)
30      End If
  
35      mvwChoixDate.Visible = False

40      Exit Sub

AfficherErreur:

45      woups "frmCédule", "mvwChoixDate_DateClick", Err, Erl
End Sub

Private Sub mvwChoixDate_LostFocus()

5       On Error GoTo AfficherErreur
 
10      mvwChoixDate.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmCédule", "mvwChoixDate_LostFocus", Err, Erl
End Sub

Private Sub mskHeure_GotFocus()

5       On Error GoTo AfficherErreur
        
        'Format d'heure
10      mskHeure.mask = "##:##"

15      Exit Sub

AfficherErreur:

20      woups "frmCédule", "mskHeure_GotFocus", Err, Erl
End Sub

Private Sub mskHeure_LostFocus()

5       On Error GoTo AfficherErreur

        'Enlève le mask
10      mskHeure.mask = vbNullString
  
        'Vide le champs si l'utilisateur n'a rien écrit
15      If mskHeure.Text = "__:__" Then
20        mskHeure.Text = vbNullString
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmCédule", "mskHeure_LostFocus", Err, Erl
End Sub

Private Sub AfficherDate()

5       On Error GoTo AfficherErreur

        'Affiche horaire de la journée et de la semaine
        'dépendant la sélection dans le calendrier
10      Dim iCompteur As Integer
  
        'Date choisie
15      m_datDateChoisie = DateSerial(mvwSelection.Year, mvwSelection.Month, mvwSelection.Day)

        'Affiche horaire jour et semaine
20      Call RemplirListerJour
25      Call RemplirListerSemaine

        'Sélectionne jour de la semaine
30      For iCompteur = 1 To 7
35        If lstjoursemaine(iCompteur).Tag = m_datDateChoisie Then
40          lstjoursemaine(iCompteur).BackColor = &HE0E0E0
45        Else
50          lstjoursemaine(iCompteur).BackColor = &HFFFFFF
55        End If
60      Next

        'Affiche cédule une journée
65      fraliste.Visible = True
70      fraAlarme.Visible = False
75      frajour.Visible = False

80      Exit Sub

AfficherErreur:

85      woups "frmCédule", "AfficherDate", Err, Erl
End Sub

Private Sub mvwSelection_GotFocus()

5       On Error GoTo AfficherErreur

10      m_bMonthViewHasFocus = True

15      Exit Sub

AfficherErreur:

20      woups "frmCédule", "mvwSelection_GotFocus", Err, Erl
End Sub

Private Sub mvwSelection_SelChange(ByVal StartDate As Date, ByVal EndDate As Date, Cancel As Boolean)

5       On Error GoTo AfficherErreur

10      If Month(m_datDateChoisie) <> mvwSelection.Month Or _
          Year(m_datDateChoisie) <> mvwSelection.Year Or _
          Day(m_datDateChoisie) <> mvwSelection.Day Then
15        Call AfficherDate
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmCédule", "mvwSelection_SelChange", Err, Erl
End Sub
