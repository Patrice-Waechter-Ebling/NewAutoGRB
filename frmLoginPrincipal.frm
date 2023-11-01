VERSION 5.00
Begin VB.Form frmLoginPrincipal 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLoginPrincipal.frx":0000
   ScaleHeight     =   2640
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
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
   Begin VB.CommandButton cmdOK 
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
   Begin VB.TextBox txtlogin 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   960
      Width           =   2535
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
End
Attribute VB_Name = "frmLoginPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

5       On Error GoTo AfficherErreur

10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmLoginPrincipal", "cmdCancel_Click", Err, Erl
End Sub

Private Sub cmdOK_Click()

5       On Error GoTo AfficherErreur

10      Dim rstEmploye As ADODB.Recordset

        'Ouvre la table
15      Set rstEmploye = New ADODB.Recordset

20      Call rstEmploye.Open("SELECT * FROM GRB_Employés WHERE loginname = '" & txtlogin.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

        'Si trouve utilisateur
25      If Not rstEmploye.EOF Then
          'si bon mot de passe, save user et quitte loggin
30        If UCase(rstEmploye.Fields("passwd")) = UCase(txtpasswd.Text) Then
35          Call SaveSetting("GRB", "Config", "LoginPrincipal", txtlogin.Text)

            'bon_passwd
40          g_bBonPasswd = True

            'UserID
45          g_sUserID = txtlogin.Text

            'Groupe de securité
50          g_iNoGroupe = rstEmploye.Fields("groupe")

            'Nom de l'employé
55          g_sEmploye = rstEmploye.Fields("employe")

60          Call InitialiserVariablesGroupe

            'Initiale de l'employé
65          g_sInitiale = rstEmploye.Fields("initiale")

70          Call rstEmploye.Close
75          Set rstEmploye = Nothing

80          Screen.MousePointer = vbHourglass

            'Fermeture du login
85          Call Unload(Me)

            'Ouverture de Dispatch
90          Call OuvrirForm(FrmDispatch, False)
95        Else
100         Call MsgBox("Mot de passe invalide!")
105       End If
110     Else
115       Call MsgBox("L'utilisateur n'existe pas!")
120     End If
        
        
125         Exit Sub

AfficherErreur:

130       woups "frmLoginPrincipal", "cmdOK_Click", Err, Erl

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

40      woups "frmLoginPrincipal", "Form_Activate", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

10      Dim rstConfig        As ADODB.Recordset
15      Dim sVersion         As String
20      Dim sDerniereVersion As String

25      If OuvrirConnection = True Then
30        sVersion = App.Major & "." & Right$("0" & App.Minor, 2) & "." & Right$("0" & App.Revision, 4)

35        Set rstConfig = New ADODB.Recordset

40        Call rstConfig.Open("SELECT DerniereVersion FROM GRB_Config", g_connData, adOpenDynamic, adLockOptimistic)

45        If Not IsNull(rstConfig.Fields("DerniereVersion")) Then
50          If rstConfig.Fields("DerniereVersion") <> "" Then
55            sDerniereVersion = rstConfig.Fields("DerniereVersion")
60          Else
65            sDerniereVersion = ""
70          End If
75        Else
80          sDerniereVersion = ""
85        End If

90        Call rstConfig.Close
95        Set rstConfig = Nothing

100       If sDerniereVersion <> sVersion Then
105         Call MsgBox("Votre version n'est pas à jour, elle sera installée!", vbOKOnly)

110         Call ShellExecute(Me.hwnd, vbNullString, App.Path & "\InstallGRB.exe", vbNullString, "C:\", SW_SHOWNORMAL)

115         Call FermerConnection

120         End
125       End If

130       txtlogin.Text = GetSetting("GRB", "Config", "LoginPrincipal", "")
135     End If

140     Exit Sub

AfficherErreur:

145     woups "frmLoginPrincipal", "Form_Load", Err, Erl
End Sub

Private Sub txtlogin_GotFocus()

5       On Error GoTo AfficherErreur

10      If txtlogin.Text <> "" Then
15        txtlogin.SelStart = 0
20        txtlogin.SelLength = Len(txtlogin.Text)
25      End If

30      Exit Sub

AfficherErreur:

35      woups "frmLoginPrincipal", "txtlogin_GotFocus", Err, Erl
End Sub
