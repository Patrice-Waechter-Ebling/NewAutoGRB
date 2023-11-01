VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLoginAdmin 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2940
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5460
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLoginAdmin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1737.049
   ScaleMode       =   0  'User
   ScaleWidth      =   5126.645
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2685
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   "Procédure de sécurite NIS8"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   2250
      TabIndex        =   2
      Top             =   975
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "S'identifier"
      Default         =   -1  'True
      Height          =   390
      Left            =   2400
      TabIndex        =   5
      Top             =   2040
      Width           =   1380
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   390
      Left            =   3840
      TabIndex        =   6
      Top             =   2040
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2250
      PasswordChar    =   "$"
      TabIndex        =   4
      Top             =   1365
      Width           =   2325
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cette option requiere un privilege de niveau 99"
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   720
      TabIndex        =   7
      Top             =   360
      Width           =   4395
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      ForeColor       =   &H00FFC0C0&
      Height          =   270
      Index           =   0
      Left            =   1065
      TabIndex        =   1
      Top             =   990
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      ForeColor       =   &H00FFC0C0&
      Height          =   270
      Index           =   1
      Left            =   1065
      TabIndex        =   3
      Top             =   1380
      Width           =   1080
   End
End
Attribute VB_Name = "frmLoginAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    If txtPassword = "Password01$" Then
        If txtUserName = "Administrateur" Or txtUserName = frmLogin.List1.List(frmLogin.Combo1.ListIndex) Then 'prevoit un mode de recuperation en utilisant une combinaison de 2 mots de passe presise
            LoginSucceeded = True
            Me.Hide
        End If
    Else
        MsgBox "Invalid Password, try again!", , Titre + " Élévation de privileges"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub

Private Sub Form_Load()
Me.Caption = "Élévation de privilege pour " + Conteneur.StatusBar1.Panels(2).Text
End Sub

