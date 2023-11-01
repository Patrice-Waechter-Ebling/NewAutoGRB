VERSION 5.00
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NewAutoGRB - Identification"
   ClientHeight    =   3090
   ClientLeft      =   1050
   ClientTop       =   1395
   ClientWidth     =   6045
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1080
      TabIndex        =   7
      Top             =   2280
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2640
      TabIndex        =   6
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Abandonner"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connexion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox motdepass 
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   2640
      PasswordChar    =   "#"
      TabIndex        =   3
      Text            =   "waechter"
      Top             =   1320
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   390
      Left            =   2640
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   780
      Width           =   2535
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mot de passe"
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   960
      TabIndex        =   2
      Top             =   1320
      Width           =   1230
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nom d'utilisateur"
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   1620
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    noyau.AquerirEmployes
    noyau.Login Combo1.List(Combo1.ListIndex - 1), motdepass.Text
    Unload Me
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
If motdepass.Text = List1.List(Combo1.ListIndex) Then
Conteneur.Show
Else
    MsgBox Err.Description + vbCrLf + Err.Source, vbCritical, Me.Caption
End If
End Sub
Private Sub Form_Load()
On Error GoTo Oups
Me.Picture = LoadPicture(App.Path + "\drapeauGRB.jpg")
    Dim g_connData As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim fld As ADODB.Field, alignment As Integer
    Dim recCount As Long, i As Long, fldName As String
    g_connData.Open "Driver={SQL Server};Server=TOUR-PATRICE\SQLEXPRESS;Database=WebGRB;Trusted_Connection=Yes;"
    rs.Open "GrbEmploye", g_connData, adOpenForwardOnly, adLockReadOnly
    Combo1.Clear
    rs.MoveFirst
    Do Until rs.EOF
        recCount = recCount + 1
        Combo1.AddItem rs.Fields("employe")
        If (rs.Fields("Actif") = True) Then Conteneur.Toolbar1.Buttons(1).ButtonMenus.Add , rs.Fields("loginname"), rs.Fields("employe")
        List1.AddItem rs.Fields("passwd")
        If recCount = MaxRecords Then Exit Do
        rs.MoveNext
    Loop
    Combo1.ListIndex = recount - 1

    Exit Sub
Oups:
    MsgBox Err.Description + vbCrLf + Err.Source, vbCritical, Me.Caption
End Sub

Private Sub motdepass_Change()
If Len(Me.motdepass.Text) > 3 Then
Command1.Enabled = True
Else
Command1.Enabled = False
End If
End Sub
