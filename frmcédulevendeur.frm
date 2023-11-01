VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmCédulevendeur 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cédule des vendeurs"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11820
   Icon            =   "frmcédulevendeur.frx":0000
   LinkTopic       =   "frmcédule"
   MaxButton       =   0   'False
   Picture         =   "frmcédulevendeur.frx":0442
   ScaleHeight     =   7110
   ScaleWidth      =   11820
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frasemaine 
      BackColor       =   &H00404040&
      Height          =   3015
      Left            =   0
      TabIndex        =   14
      Top             =   4080
      Width           =   11805
      Begin MSComctlLib.ListView lstjoursemaine 
         Height          =   2220
         Index           =   1
         Left            =   0
         TabIndex        =   15
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   3916
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
            Text            =   "Heure"
            Object.Width           =   741
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "nom"
            Object.Width           =   2152
         EndProperty
      End
      Begin MSComctlLib.ListView lstjoursemaine 
         Height          =   2220
         Index           =   2
         Left            =   1680
         TabIndex        =   16
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   3916
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
            Text            =   "Heure"
            Object.Width           =   741
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "nom"
            Object.Width           =   2152
         EndProperty
      End
      Begin MSComctlLib.ListView lstjoursemaine 
         Height          =   2220
         Index           =   3
         Left            =   3360
         TabIndex        =   17
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   3916
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
            Text            =   "Heure"
            Object.Width           =   741
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "nom"
            Object.Width           =   2152
         EndProperty
      End
      Begin MSComctlLib.ListView lstjoursemaine 
         Height          =   2220
         Index           =   4
         Left            =   5040
         TabIndex        =   18
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   3916
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
            Text            =   "Heure"
            Object.Width           =   741
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "nom"
            Object.Width           =   2152
         EndProperty
      End
      Begin MSComctlLib.ListView lstjoursemaine 
         Height          =   2220
         Index           =   5
         Left            =   6720
         TabIndex        =   19
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   3916
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
            Text            =   "Heure"
            Object.Width           =   741
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "nom"
            Object.Width           =   2152
         EndProperty
      End
      Begin MSComctlLib.ListView lstjoursemaine 
         Height          =   2220
         Index           =   6
         Left            =   8400
         TabIndex        =   20
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   3916
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
            Text            =   "Heure"
            Object.Width           =   741
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "nom"
            Object.Width           =   2152
         EndProperty
      End
      Begin MSComctlLib.ListView lstjoursemaine 
         Height          =   2220
         Index           =   7
         Left            =   10080
         TabIndex        =   21
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   3916
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
            Text            =   "Heure"
            Object.Width           =   741
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "nom"
            Object.Width           =   2152
         EndProperty
      End
      Begin VB.Label D 
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
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label6 
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
         Left            =   1800
         TabIndex        =   34
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label7 
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
         Left            =   3480
         TabIndex        =   33
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label8 
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
         Left            =   5160
         TabIndex        =   32
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label9 
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
         Left            =   6840
         TabIndex        =   31
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label10 
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
         Left            =   8520
         TabIndex        =   30
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label12 
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
         Left            =   10200
         TabIndex        =   29
         Top             =   360
         Width           =   615
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   3360
         X2              =   3360
         Y1              =   120
         Y2              =   2280
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   5040
         X2              =   5040
         Y1              =   120
         Y2              =   2280
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   6720
         X2              =   6720
         Y1              =   120
         Y2              =   2280
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         X1              =   8400
         X2              =   8400
         Y1              =   120
         Y2              =   2280
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         X1              =   10080
         X2              =   10080
         Y1              =   120
         Y2              =   2280
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   1680
         X2              =   1680
         Y1              =   120
         Y2              =   2520
      End
      Begin VB.Label lbljour 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   28
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lbljour 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   27
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lbljour 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   3960
         TabIndex        =   26
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lbljour 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   5640
         TabIndex        =   25
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lbljour 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   7320
         TabIndex        =   24
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lbljour 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   9000
         TabIndex        =   23
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lbljour 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   10680
         TabIndex        =   22
         Top             =   360
         Width           =   495
      End
   End
   Begin MSACAL.Calendar Calemploye 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   5175
      _Version        =   524288
      _ExtentX        =   9128
      _ExtentY        =   5741
      _StockProps     =   1
      BackColor       =   0
      Year            =   2002
      Month           =   8
      Day             =   2
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   16777215
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   16777215
      GridLinesColor  =   8421504
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   16777215
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraliste 
      BackColor       =   &H00C0C0C0&
      Height          =   3855
      Left            =   5280
      TabIndex        =   7
      Top             =   120
      Width           =   6495
      Begin VB.CommandButton cmdajouter 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ajouter"
         Height          =   495
         Left            =   3360
         TabIndex        =   12
         Top             =   3300
         Width           =   1455
      End
      Begin VB.CommandButton cmdsupprimer 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Supprimer"
         Height          =   495
         Left            =   4920
         TabIndex        =   11
         Top             =   3300
         Width           =   1455
      End
      Begin MSComctlLib.ListView Lstjour 
         Height          =   2895
         Left            =   0
         TabIndex        =   10
         Top             =   360
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   5106
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
            Text            =   "Heure"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "À téléphoner"
            Object.Width           =   5345
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Suivi chez le client"
            Object.Width           =   5345
         EndProperty
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Heure"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   36
         Top             =   120
         Width           =   600
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Suivi chez le client"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3630
         TabIndex        =   9
         Top             =   120
         Width           =   3030
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "À téléphoner"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   120
         Width           =   3030
      End
   End
   Begin VB.Frame frajour 
      Height          =   3855
      Left            =   5280
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CheckBox chkAlarme 
         Alignment       =   1  'Right Justify
         Caption         =   "Alarme"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   39
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtclient 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   2040
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   2160
         Width           =   4335
      End
      Begin VB.CommandButton cmdannuler 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Annuler"
         Height          =   495
         Left            =   4920
         TabIndex        =   6
         Top             =   3240
         Width           =   1455
      End
      Begin VB.CommandButton cmdenreg 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Enregistrer"
         Height          =   495
         Left            =   3360
         TabIndex        =   5
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox txttelephone 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   2040
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1440
         Width           =   4335
      End
      Begin MSMask.MaskEdBox mskHeure 
         Height          =   360
         Left            =   2040
         TabIndex        =   38
         Top             =   960
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label lblDate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   6135
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Heure :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   37
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Suivi chez le client :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   120
         TabIndex        =   3
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "À téléphoner :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   240
         TabIndex        =   2
         Top             =   1440
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmCédulevendeur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const I_COL_JOUR_HEURE      As Integer = 0
Private Const I_COL_JOUR_TELEPHONER As Integer = 1
Private Const I_COL_JOUR_SUIVI      As Integer = 2

Private Const I_COL_SEMAINE_HEURE   As Integer = 0
Private Const I_COL_SEMAINE_MESSAGE As Integer = 1

Private m_datDateChoisie As Date
Private m_bModeAjouter   As Boolean

Private Sub Calemploye_Click()

5       On Error GoTo AfficherErreur

        'Affiche horaire de la journée et de la semaine
        'dépendant la sélection dans le calendrier
10      Dim iCompteur As Integer

        'Date choisie
15      m_datDateChoisie = DateSerial(calEmploye.Year, calEmploye.Month, calEmploye.Day)
        
        'Affiche horaire jour et semaine
20      Call RemplirListerJour
25      Call RemplirListerSemaine

        'selectionne jour de la semaine
30      For iCompteur = 1 To 7
35        If lstjoursemaine(iCompteur).Tag = m_datDateChoisie Then
40          lstjoursemaine(iCompteur).BackColor = &HE0E0E0
45        Else
50          lstjoursemaine(iCompteur).BackColor = &HFFFFFF
55        End If
60      Next

        'affiche cedule une journee
65      fraliste.Visible = True
70      fraJour.Visible = False

75      Exit Sub

AfficherErreur:

80      Call AfficherErreur(Me, "Calemploye_Click", Err, Erl)
End Sub

Private Sub cmdAjouter_Click()

5       On Error GoTo AfficherErreur

10      Dim sMois As String
15      Dim sJour As String

        'met en mode ajouter et affiche champ pour entrer des données
20      m_bModeAjouter = True
25      fraliste.Visible = False
30      fraJour.Visible = True

35      chkAlarme.Visible = True

40      chkAlarme.Value = vbUnchecked
        
        'vide champ text
45      mskHeure.Text = ""
50      txtClient.Text = vbNullString
55      txtTelephone.Text = vbNullString

60      Select Case Month(m_datDateChoisie)
          Case 1:  sMois = "Janvier"
65        Case 2:  sMois = "Février"
70        Case 3:  sMois = "Mars"
75        Case 4:  sMois = "Avril"
80        Case 5:  sMois = "Mai"
85        Case 6:  sMois = "Juin"
90        Case 7:  sMois = "Juillet"
95        Case 8:  sMois = "Août"
100       Case 9:  sMois = "Septembre"
105       Case 10: sMois = "Octobre"
110       Case 11: sMois = "Novembre"
115       Case 12: sMois = "Décembre"
120     End Select

125     Select Case Weekday(m_datDateChoisie)
          Case 1: sJour = "Dimanche"
130       Case 2: sJour = "Lundi"
135       Case 3: sJour = "Mardi"
140       Case 4: sJour = "Mercredi"
145       Case 5: sJour = "Jeudi"
150       Case 6: sJour = "Vendredi"
155       Case 7: sJour = "Samedi"
160     End Select

165     lblDate.Caption = sJour & ", le " & Day(m_datDateChoisie) & " " & sMois & " " & Year(m_datDateChoisie)

170     Exit Sub

AfficherErreur:

175     Call AfficherErreur(Me, "cmdAjouter_Click", Err, Erl)
End Sub

Private Sub cmdAnnuler_Click()

5       On Error GoTo AfficherErreur

        'quitte ecran pour ajouté ou modifié
10      fraliste.Visible = True
15      fraJour.Visible = False

20      Exit Sub

AfficherErreur:

25      Call AfficherErreur(Me, "cmdAnnuler_Click", Err, Erl)
End Sub

Private Sub cmdenreg_Click()

5       On Error GoTo AfficherErreur
       
        'Enregistre
10      Dim rstCédule    As ADODB.Recordset
15      Dim rstAlarme    As ADODB.Recordset
20      Dim rstEmploye   As ADODB.Recordset
25      Dim iNoEmploye   As Integer
30      Dim sMsgAlarme   As String
35      Dim bModifAlarme As Boolean

        'Ouvre la table
40      If m_bModeAjouter = True Then
45        Set rstCédule = New ADODB.Recordset

50        Call rstCédule.Open("SELECT * FROM GRB_cédulevendeur", g_connData, adOpenDynamic, adLockOptimistic)
      
55        Call rstCédule.AddNew
60      Else
65        If Lstjour.SelectedItem.ListSubItems(I_COL_JOUR_TELEPHONER).Tag = "A" Then
70          Set rstAlarme = New ADODB.Recordset

75          Call rstAlarme.Open("SELECT * FROM GRB_Alarmes WHERE IDAlarme = " & Lstjour.SelectedItem.Tag, g_connData, adOpenDynamic, adLockOptimistic)

80          bModifAlarme = True
85        Else
90          Set rstCédule = New ADODB.Recordset

95          Call rstCédule.Open("SELECT * FROM GRB_CéduleVendeur WHERE noenreg = " & Lstjour.SelectedItem.Tag, g_connData, adOpenDynamic, adLockOptimistic)
100       End If
105     End If
    
110     If bModifAlarme = True Then
115       rstAlarme.Fields("Heure") = mskHeure.Text

120       sMsgAlarme = "À téléphoner : " & txtTelephone.Text
125       sMsgAlarme = sMsgAlarme & vbNewLine & "Suivi chez le client : " & txtClient.Text

130       rstAlarme.Fields("Message") = sMsgAlarme

135       Call rstAlarme.Update

140       Call rstAlarme.Close
145       Set rstAlarme = Nothing
150     Else
155       rstCédule.Fields("Date_cedulé") = ConvertDate(m_datDateChoisie)

160       rstCédule.Fields("Heure") = mskHeure.Text
        
165       If txtTelephone.Text = vbNullString Then
170         rstCédule.Fields("a_telephoner") = " "
175       Else
180         rstCédule.Fields("a_telephoner") = txtTelephone.Text
185       End If
        
190       rstCédule.Fields("JourSemaine") = Weekday(m_datDateChoisie)
  
195       If txtClient.Text = vbNullString Then
200         rstCédule.Fields("Client") = " "
205       Else
210         rstCédule.Fields("Client") = txtClient.Text
215       End If
       
220       Call rstCédule.Update

225       Set rstEmploye = New ADODB.Recordset

230       Call rstEmploye.Open("SELECT * FROM GRB_Employés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)

235       iNoEmploye = rstEmploye.Fields("NoEmploye")

240       Call rstEmploye.Close
245       Set rstEmploye = Nothing
        
          'Enregistrement des alarmes
250       Set rstAlarme = New ADODB.Recordset
          
255       Call rstAlarme.Open("SELECT * FROM GRB_Alarmes WHERE IDCédule = " & rstCédule.Fields("NoEnreg") & " AND NoEmploye = " & iNoEmploye, g_connData, adOpenDynamic, adLockOptimistic)

260       If chkAlarme.Value = vbChecked Then
265         If rstAlarme.EOF Then
270           Call rstAlarme.AddNew

275           rstAlarme.Fields("NoEmploye") = iNoEmploye
280           rstAlarme.Fields("IDCédule") = rstCédule.Fields("NoEnreg")
285         End If

290         rstAlarme.Fields("TypeCédule") = "CV"
295         rstAlarme.Fields("Date") = ConvertDate(m_datDateChoisie)
300         rstAlarme.Fields("Heure") = mskHeure.Text

305         sMsgAlarme = "À téléphoner : " & txtTelephone.Text
310         sMsgAlarme = sMsgAlarme & vbNewLine & "Suivi chez le client : " & txtClient.Text

315         rstAlarme.Fields("Message") = sMsgAlarme

320         rstAlarme.Fields("JourSemaine") = Weekday(m_datDateChoisie)

325         Call rstAlarme.Update
330       Else
335         If Not rstAlarme.EOF Then
340           Call rstAlarme.Delete
345         End If
350       End If

355       Call rstAlarme.Close
360       Set rstAlarme = Nothing

365       Call rstCédule.Close
370       Set rstCédule = Nothing
375     End If
            
        'Quitte l'écran pour ajouter ou modifier
380     fraliste.Visible = True
385     fraJour.Visible = False
       
390     Call RemplirListerJour
395     Call RemplirListerSemaine
    
400     m_bModeAjouter = False

405     Exit Sub

AfficherErreur:

410     Call AfficherErreur(Me, "cmdenreg_Click", Err, Erl)
End Sub

Private Sub cmdSupprimer_Click()

5       On Error GoTo AfficherErreur

10      Dim lRecordID As Long

15      If Lstjour.ListItems.Count > 0 Then
20        If MsgBox("Voulez-vous supprimer cette enregistrement?", vbYesNo) = vbYes Then
25          lRecordID = Lstjour.SelectedItem.Tag

30          If Lstjour.SelectedItem.ListSubItems(I_COL_JOUR_TELEPHONER).Tag = "A" Then
35            Call g_connData.Execute("DELETE * FROM GRB_Alarmes WHERE IDAlarme = " & lRecordID)
40          Else
45            Call g_connData.Execute("DELETE * FROM GRB_cédulevendeur WHERE noenreg = " & lRecordID)
50          End If
55        End If
60      End If
        
        'Mise a jour des lister
65      Call RemplirListerJour
70      Call RemplirListerSemaine

75      Exit Sub

AfficherErreur:

80      Call AfficherErreur(Me, "cmdSupprimer_Click", Err, Erl)
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer

15      g_bCeduleVendeursOuverte = True
        
        'Met à jour l'écran
20      calEmploye.Month = Month(Date)
25      calEmploye.Year = Year(Date)
30      calEmploye.Day = Day(Date)

35      m_datDateChoisie = Date

        'Rempli les lister
40      Call RemplirListerJour
45      Call RemplirListerSemaine

        'Sélectionne jour de la semaine
50      For iCompteur = 1 To 7
55        If lstjoursemaine(iCompteur).Tag = m_datDateChoisie Then
60          lstjoursemaine(iCompteur).BackColor = &HE0E0E0
65        Else
70          lstjoursemaine(iCompteur).BackColor = &HFFFFFF
75        End If
80      Next

85      Exit Sub

AfficherErreur:

90      Call AfficherErreur(Me, "Form_Load", Err, Erl)
End Sub

Public Sub RemplirListerJour()

5       On Error GoTo AfficherErreur

        'Remplis lister une journée
10      Dim rstCédule As ADODB.Recordset
15      Dim itmCedule As ListItem
  
        'Vide le lister
20      Call Lstjour.ListItems.Clear
  
25      Set rstCédule = New ADODB.Recordset
  
30      Call rstCédule.Open("SELECT * FROM GRB_céduleVendeur WHERE date_cedulé = '" & ConvertDate(m_datDateChoisie) & "'", g_connData, adOpenDynamic, adLockOptimistic)
    
        'Tant il y a de employé cedulé , ajoute dans lister
35      Do While Not rstCédule.EOF
40        Set itmCedule = Lstjour.ListItems.Add
            
45        itmCedule.Tag = rstCédule.Fields("noenreg")

50        If Not IsNull(rstCédule.Fields("Heure")) Then
55          itmCedule.Text = rstCédule.Fields("Heure")
60        Else
65          itmCedule.Text = ""
70        End If
            
75        If IsNull(rstCédule.Fields("a_telephoner")) Then
80          itmCedule.SubItems(I_COL_JOUR_TELEPHONER) = vbNullString
85        Else
90          itmCedule.SubItems(I_COL_JOUR_TELEPHONER) = rstCédule.Fields("a_telephoner")
95        End If

100       itmCedule.ListSubItems(I_COL_JOUR_TELEPHONER).Tag = "C"
           
105       If IsNull(rstCédule.Fields("Client")) Then
110         itmCedule.SubItems(I_COL_JOUR_SUIVI) = vbNullString
115       Else
120         itmCedule.SubItems(I_COL_JOUR_SUIVI) = rstCédule.Fields("Client")
125       End If
            
130       Call rstCédule.MoveNext
135     Loop
      
140     Call rstCédule.Close
145     Set rstCédule = Nothing

150     Call RemplirListerJourAlarme

155     Exit Sub

AfficherErreur:

160     Call AfficherErreur(Me, "RemplirListerJour", Err, Erl)
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
45      Dim sSubItems       As String

        'Vide les ListView
50      For iCompteur = 1 To 7
55        Call lstjoursemaine(iCompteur).ListItems.Clear

60        lstjoursemaine(iCompteur).Sorted = False
65      Next
    
70      iJourSemaine = Weekday(m_datDateChoisie)
75      datPremiereDate = m_datDateChoisie
80      datDerniereDate = m_datDateChoisie
    
        'Trouve première date de la semaine
85      Do While Not Weekday(datPremiereDate) = 1
90        datPremiereDate = datPremiereDate - 1
95      Loop
    
        'Trouve dernière date de la semaine
100     Do While Not Weekday(datDerniereDate) = 7
105       datDerniereDate = datDerniereDate + 1
110     Loop
    
        'Sélectionne la semaine courante
115     Set rstCédule = New ADODB.Recordset
        
120     Call rstCédule.Open("SELECT * FROM GRB_cédulevendeur WHERE cdate(date_cedulé) <= cdate('" & CStr(datDerniereDate) & "') AND cdate(date_cedulé) >= cdate('" & CStr(datPremiereDate) & "')", g_connData, adOpenDynamic, adLockOptimistic)
    
125     For iCompteur = 1 To 7
          'Pour écrire le jour
130       lblJour(iCompteur - 1).Caption = Day(datPremiereDate + iCompteur - 1)
      
          'Garde en memoire la date des ListView
135       lstjoursemaine(iCompteur).Tag = datPremiereDate + iCompteur - 1
140     Next
                
145     Do While Not rstCédule.EOF
          'Ajoute dans le ListView, dépendant le jour de la semaine
150       Set itmSemaine = lstjoursemaine(rstCédule.Fields("joursemaine")).ListItems.Add
      
155       itmSemaine.Tag = rstCédule.Fields("noenreg")
            
160       If Not IsNull(rstCédule.Fields("Heure")) Then
165         itmSemaine.Text = rstCédule.Fields("Heure")
170       Else
175         itmSemaine.Text = ""
180       End If
            
185       sSubItems = vbNullString
            
190       If Trim(rstCédule.Fields("Client")) <> vbNullString Then
195         sSubItems = rstCédule.Fields("Client")
200       End If
       
205       If Trim(rstCédule.Fields("a_telephoner")) <> vbNullString Then
210         sSubItems = sSubItems & " " & rstCédule.Fields("a_telephoner")
215       End If
               
220       itmSemaine.SubItems(I_COL_SEMAINE_MESSAGE) = sSubItems
        
225       Call rstCédule.MoveNext
230     Loop
     
235     Call rstCédule.Close
240     Set rstCédule = Nothing

245     Call RemplirListerSemaineAlarme

250     Exit Sub

AfficherErreur:

255     Call AfficherErreur(Me, "RemplirListerSemaine", Err, Erl)
End Sub

Private Sub Form_Unload(Cancel As Integer)

5       On Error GoTo AfficherErreur

10      g_bCeduleVendeursOuverte = False

15      Exit Sub

AfficherErreur:

20      Call AfficherErreur(Me, "Form_Unload", Err, Erl)
End Sub

Private Sub Lstjour_DblClick()

5       On Error GoTo AfficherErreur

10      If Lstjour.ListItems.Count > 0 Then

15        If Lstjour.SelectedItem.ListSubItems(I_COL_JOUR_TELEPHONER).Tag = "C" Then
20          chkAlarme.Visible = True
25          chkAlarme.Value = vbUnchecked
30        Else
35          chkAlarme.Visible = False
40        End If

          'Affiche en mode modification
45        m_bModeAjouter = False
50        fraliste.Visible = False
55        fraJour.Visible = True

60        mskHeure.Text = Lstjour.SelectedItem.Text
65        txtTelephone.Text = Lstjour.SelectedItem.SubItems(I_COL_JOUR_TELEPHONER)
70        txtClient.Text = Lstjour.SelectedItem.SubItems(I_COL_JOUR_SUIVI)
75      End If

80      Exit Sub

AfficherErreur:

85      Call AfficherErreur(Me, "Lstjour_DblClick", Err, Erl)
End Sub

Private Sub Lstjour_KeyDown(KeyCode As Integer, Shift As Integer)

5       On Error GoTo AfficherErreur

10      Dim lRecordID As Long

15      If Lstjour.ListItems.Count > 0 Then
20        If KeyCode = vbKeyDelete Then
25          If MsgBox("Voulez-vous supprimer cette enregistrement?", vbYesNo) = vbYes Then
30            lRecordID = Lstjour.SelectedItem.Tag
          
35            Call g_connData.Execute("DELETE * FROM GRB_cédulevendeur WHERE noenreg = " & lRecordID)
40          End If
        
            'Mise à jour des listers
45          Call RemplirListerJour
50          Call RemplirListerSemaine
55        End If
60      End If

65      Exit Sub

AfficherErreur:

70      Call AfficherErreur(Me, "Lstjour_KeyDown", Err, Erl)
End Sub

Private Sub lstjoursemaine_Click(Index As Integer)

5       On Error GoTo AfficherErreur

10      Dim iCompteur As Integer
      
        'Initialise la couleur en blanc
15      For iCompteur = 1 To 7
20        lstjoursemaine(iCompteur).BackColor = &HFFFFFF
25      Next

        'Sélectionne jour de semaine
30      lstjoursemaine(Index).BackColor = &HE0E0E0
  
35      Call AjusterColonne(Index)

        'Sélectionne dans calendrier
40      calEmploye.Day = Day(CStr(lstjoursemaine(Index).Tag))
45      calEmploye.Month = Month(CStr(lstjoursemaine(Index).Tag))
50      calEmploye.Year = Year(CStr(lstjoursemaine(Index).Tag))

        'Date choisie
55      m_datDateChoisie = DateSerial(calEmploye.Year, calEmploye.Month, calEmploye.Day)
  
        'Affiche horaire jour
60      Call RemplirListerJour

        'Affiche cédule une journée
65      fraliste.Visible = True
70      fraJour.Visible = False

75      Exit Sub

AfficherErreur:

80      Call AfficherErreur(Me, "lstjoursemaine_Click", Err, Erl)
End Sub

Private Sub AjusterColonne(ByVal iIndex As Integer)

5       On Error GoTo AfficherErreur

10      Dim iPlusLong As Integer
15      Dim iCompteur As Integer
20      Dim sTexte    As String
     
25      iPlusLong = 0

30      For iCompteur = 1 To lstjoursemaine(iIndex).ListItems.Count
35        If TextWidth(lstjoursemaine(iIndex).ListItems(iCompteur).SubItems(1)) > iPlusLong Then
40          iPlusLong = TextWidth(lstjoursemaine(iIndex).ListItems(iCompteur).SubItems(1))
45          sTexte = lstjoursemaine(iIndex).ListItems(iCompteur).SubItems(1)
50        End If
55      Next
          
60      lstjoursemaine(iIndex).ColumnHeaders(2).Width = iPlusLong

65      If lstjoursemaine(iIndex).ColumnHeaders(2).Width < lstjoursemaine(iIndex).Width Then
70        lstjoursemaine(iIndex).ColumnHeaders(2).Width = 1450
75      End If

80      Exit Sub

AfficherErreur:

85      Call AfficherErreur(Me, "AjusterColonne", Err, Erl)
End Sub

Private Sub lstjoursemaine_LostFocus(Index As Integer)

5       On Error GoTo AfficherErreur

10      lstjoursemaine(Index).ColumnHeaders(2).Width = 1300

15      Exit Sub

AfficherErreur:

20      Call AfficherErreur(Me, "lstjoursemaine_LostFocus", Err, Erl)
End Sub

Private Sub RemplirListerJourAlarme()

5       On Error GoTo AfficherErreur

        'Remplis lister une journée
10      Dim rstAlarme   As ADODB.Recordset
15      Dim rstEmploye  As ADODB.Recordset
20      Dim iNoEmploye  As Integer
25      Dim itmAlarme   As ListItem
30      Dim sTelephoner As String
35      Dim sSuivi      As String

40      Set rstEmploye = New ADODB.Recordset

45      Call rstEmploye.Open("SELECT * FROM GRB_Employés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)

50      iNoEmploye = rstEmploye.Fields("NoEmploye")

55      Call rstEmploye.Close
60      Set rstEmploye = Nothing

65      Set rstAlarme = New ADODB.Recordset

70      Call rstAlarme.Open("SELECT * FROM GRB_Alarmes WHERE Date = '" & ConvertDate(m_datDateChoisie) & "' AND NoEmploye = " & iNoEmploye & " AND TypeCédule = 'CV' ORDER BY Date, Heure", g_connData, adOpenDynamic, adLockOptimistic)

        'Tant il y a de employé cedulé , ajoute dans lister
75      Do While Not rstAlarme.EOF
80        Set itmAlarme = Lstjour.ListItems.Add

85        itmAlarme.Text = rstAlarme.Fields("Heure")
90        itmAlarme.Tag = rstAlarme.Fields("IDAlarme")

95        sTelephoner = Left(rstAlarme.Fields("Message"), InStr(1, rstAlarme.Fields("Message"), vbNewLine))
100       sTelephoner = Replace(sTelephoner, "À téléphoner : ", "")
105       sTelephoner = Replace(sTelephoner, vbCr, "")

110       sSuivi = Right(rstAlarme.Fields("Message"), Len(rstAlarme.Fields("Message")) - InStr(rstAlarme.Fields("Message"), vbNewLine))
115       sSuivi = Replace(sSuivi, "Suivi chez le client : ", "")
120       sSuivi = Replace(sSuivi, vbLf, "")

125       itmAlarme.SubItems(I_COL_JOUR_TELEPHONER) = sTelephoner
130       itmAlarme.SubItems(I_COL_JOUR_SUIVI) = sSuivi

135       itmAlarme.ListSubItems(I_COL_JOUR_TELEPHONER).Tag = "A"

140       Call itmAlarme.ListSubItems.Add(, , rstAlarme.Fields("Message"))

145       itmAlarme.ForeColor = COLOR_BLEU
150       itmAlarme.ListSubItems(1).ForeColor = COLOR_BLEU
155       itmAlarme.ListSubItems(2).ForeColor = COLOR_BLEU

160       Call rstAlarme.MoveNext
165     Loop

170     Call rstAlarme.Close
175     Set rstAlarme = Nothing

180     Exit Sub

AfficherErreur:

185     Call AfficherErreur(Me, "RemplirListerJourAlarme", Err, Erl)
End Sub

Private Sub RemplirListerSemaineAlarme()

5      On Error GoTo AfficherErreur

        'Remplis une semaine
10      Dim rstAlarme       As ADODB.Recordset
15      Dim rstEmploye      As ADODB.Recordset
20      Dim iNoEmploye      As Integer
25      Dim iJourSemaine    As Integer
30      Dim datPremiereDate As Date
35      Dim datDerniereDate As Date
40      Dim itmSemaine      As ListItem
45      Dim sTelephoner     As String
50      Dim sSuivi          As String

55      iJourSemaine = Weekday(m_datDateChoisie)
60      datPremiereDate = m_datDateChoisie
65      datDerniereDate = m_datDateChoisie

        'Trouve première date de la semaine
70      Do While Not Weekday(datPremiereDate) = 1
75        datPremiereDate = datPremiereDate - 1
80      Loop

        'Trouve dernière date de la semaine
85      Do While Not Weekday(datDerniereDate) = 7
90        datDerniereDate = datDerniereDate + 1
95      Loop

100     Set rstEmploye = New ADODB.Recordset

105     Call rstEmploye.Open("SELECT * FROM GRB_Employés WHERE loginname = '" & g_sUserID & "'", g_connData, adOpenDynamic, adLockOptimistic)

110     iNoEmploye = rstEmploye.Fields("NoEmploye")

115     Call rstEmploye.Close
120     Set rstEmploye = Nothing

        'Sélectionne la semaine courante
125     Set rstAlarme = New ADODB.Recordset
        
130     Call rstAlarme.Open("SELECT * FROM GRB_Alarmes WHERE Date BETWEEN '" & ConvertDate(datPremiereDate) & "' AND '" & ConvertDate(datDerniereDate) & "' AND NoEmploye = " & iNoEmploye & " AND TypeCédule = 'CV' ORDER BY Date, Heure", g_connData, adOpenDynamic, adLockOptimistic)

135     Do While Not rstAlarme.EOF
          'Ajoute dans le ListView, dépendant le jour de la semaine
140       Set itmSemaine = lstjoursemaine(rstAlarme.Fields("JourSemaine")).ListItems.Add

145       itmSemaine.Tag = rstAlarme.Fields("IDAlarme")

150       itmSemaine.Text = rstAlarme.Fields("Heure")

155       sTelephoner = Left(rstAlarme.Fields("Message"), InStr(1, rstAlarme.Fields("Message"), vbNewLine))
160       sTelephoner = Replace(sTelephoner, "À téléphoner : ", "")
165       sTelephoner = Replace(sTelephoner, vbCr, "")

170       sSuivi = Right(rstAlarme.Fields("Message"), Len(rstAlarme.Fields("Message")) - InStr(rstAlarme.Fields("Message"), vbNewLine))
175       sSuivi = Replace(sSuivi, "Suivi chez le client : ", "")
180       sSuivi = Replace(sSuivi, vbLf, "")

185       itmSemaine.SubItems(I_COL_SEMAINE_MESSAGE) = sSuivi & " " & sTelephoner
          
          'Met en noir
190       itmSemaine.ForeColor = COLOR_BLEU
195       itmSemaine.ListSubItems(I_COL_SEMAINE_MESSAGE).ForeColor = COLOR_BLEU

200       Call rstAlarme.MoveNext
205     Loop

210     Call rstAlarme.Close
215     Set rstAlarme = Nothing

220     Exit Sub

AfficherErreur:

225     Call AfficherErreur(Me, "RemplirListerSemaineAlarme", Err, Erl)
End Sub

Private Sub mskHeure_GotFocus()

5       On Error GoTo AfficherErreur
        
        'Format d'heure
10      mskHeure.Mask = "##:##"

15      Exit Sub

AfficherErreur:

20      Call AfficherErreur(Me, "mskHeure_GotFocus", Err, Erl)
End Sub

Private Sub mskHeure_LostFocus()

5       On Error GoTo AfficherErreur

        'Enlève le mask
10      mskHeure.Mask = vbNullString
  
        'Vide le champs si l'utilisateur n'a rien écrit
15      If mskHeure.Text = "__:__" Then
20        mskHeure.Text = vbNullString
25      End If

30      Exit Sub

AfficherErreur:

35      Call AfficherErreur(Me, "mskHeure_LostFocus", Err, Erl)
End Sub
