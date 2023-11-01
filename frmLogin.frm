VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   2670
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtlogin 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox txtpasswd 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1440
      Width           =   2535
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00C0C0C0&
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   5
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Utilisateur:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Mot de passe:"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_frmSource As Form

Private Sub cmdcancel_Click()

5       On Error GoTo AfficherErreur

10      g_bBonPasswd = False
  
        'ferme le login
15      Call Unload(Me)

20      Exit Sub

AfficherErreur:

25      Call AfficherErreur("frmLogin", "cmdcancel_Click", Err, Erl)
End Sub

Private Sub cmdOK_Click()

5       On Error GoTo AfficherErreur

10      Dim rstEmploye As ADODB.Recordset

        'Ouverture de la table
15      Set rstEmploye = New ADODB.Recordset
        
20      Call rstEmploye.Open("SELECT * FROM GRB_employés WHERE loginname = '" & txtlogin.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)
    
        'Si trouve utilisateur
25      If Not rstEmploye.EOF Then
          'si bon mot de passe, save user et quitte loggin
30        If UCase(rstEmploye.Fields("passwd")) = UCase(txtpasswd.Text) Then
35          g_bBonPasswd = True

40          Call SaveSetting("GRB", "Config", "LoginPunch", txtlogin.Text)
        
45          m_frmSource.m_iNoGroupe = rstEmploye.Fields("Groupe")
        
50          m_frmSource.m_sUserID = txtlogin.Text
      
55          Call Unload(Me)
60        Else
65          Call MsgBox("Mot de passe invalide!")
70        End If
75      Else
80        Call MsgBox("L'utilisateur n'existe pas!")
85      End If
    
90      Call rstEmploye.Close
95      Set rstEmploye = Nothing

100     Exit Sub

AfficherErreur:

105      Call AfficherErreur("frmLogin", "cmdOK_Click", Err, Erl)
End Sub

Public Sub Afficher(ByVal frmSource As Form)

5       On Error GoTo AfficherErreur

10      Set m_frmSource = frmSource

15      Call Me.Show(vbModal)

20      Exit Sub

AfficherErreur:

25      Call AfficherErreur("frmLogin", "Afficher", Err, Erl)
End Sub

Private Sub Form_Activate()
    
5       On Error GoTo AfficherErreur

10      If txtlogin.Text = "" Then
15        Call txtlogin.SetFocus
20      Else
25        Call txtpasswd.SetFocus
30      End If

35      Exit Sub

AfficherErreur:

40      Call AfficherErreur("frmLogin", "Form_Activate", Err, Erl)
End Sub

Private Sub Form_Load()
  
5       On Error GoTo AfficherErreur

10      txtlogin.Text = GetSetting("GRB", "Config", "LoginPunch", "")

20      Exit Sub

AfficherErreur:

25      Call AfficherErreur("frmLogin", "Form_Load", Err, Erl)
End Sub

Private Sub txtlogin_GotFocus()

5       On Error GoTo AfficherErreur

10      If txtlogin.Text <> "" Then
15        txtlogin.SelStart = 0
20        txtlogin.SelLength = Len(txtlogin.Text)
25      End If

30      Exit Sub

AfficherErreur:

35      Call AfficherErreur("frmLogin", "txtlogin_GotFocus", Err, Erl)
End Sub

