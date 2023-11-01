VERSION 5.00
Begin VB.Form frmLoginPunch 
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
   Moveable        =   0   'False
   Picture         =   "frmLoginPunch.frx":0000
   ScaleHeight     =   2670
   ScaleWidth      =   4665
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtlogin 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox txtpasswd 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   1
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
      TabIndex        =   2
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
      TabIndex        =   3
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
      TabIndex        =   5
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
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "frmLoginPunch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdcancel_Click()

10      On Error GoTo AfficherErreur

20      g_bBonPasswd = False
  
        'ferme le login
30      Call Unload(Me)

40      Exit Sub

AfficherErreur:

50      Call AfficherErreur(Me, "cmdcancel_Click", Err, Erl)
End Sub

Private Sub cmdOK_Click()

10      On Error GoTo AfficherErreur

20      Dim rstEmploye As ADODB.Recordset

        'Ouverture de la table
30      Set rstEmploye = OuvrirRecordset("*", "GRB_employés", "loginname = '" & txtlogin.Text & "'", "")
    
        'Si trouve utilisateur
40      If Not rstEmploye.EOF Then
          'si bon mot de passe, save user et quitte loggin
50        If UCase(rstEmploye.Fields("passwd")) = UCase(txtpasswd.Text) Then
60          g_bBonPasswd = True
        
70          frmChoixPunch.m_sUserID = txtlogin.Text
      
80          Call Unload(Me)
90        Else
100         Call MsgBox("Mot de passe invalide!")
110       End If
120     Else
130       Call MsgBox("L'utilisateur n'existe pas!")
140     End If
    
150     Call rstEmploye.Close
160     Set rstEmploye = Nothing

170     Exit Sub

AfficherErreur:

180     Call AfficherErreur(Me, "cmdOK_Click", Err, Erl)
End Sub
