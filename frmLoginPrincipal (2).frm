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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   4680
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

 On Error GoTo Oups

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmLoginPrincipal", "cmdCancel_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdOK_Click()

 On Error GoTo Oups

 Dim rstEmploye As ADODB.Recordset

 'Ouvre la table
 Set rstEmploye = New ADODB.Recordset

 Call rstEmploye.Open("SELECT * FROM GrbEmployés WHERE loginname = '" & txtlogin.Text & "'", g_connData, adOpenDynamic, adLockOptimistic)

 'Si trouve utilisateur
 If Not rstEmploye.EOF Then
 'si bon mot de passe, save user et quitte loggin
 If UCase(rstEmploye.Fields("passwd")) = UCase(txtpasswd.Text) Then
 Call SaveSetting("GRB", "Config", "LoginPrincipal", txtlogin.Text)

 'bon_passwd
 g_bBonPasswd = True

 'UserID
 g_sUserID = txtlogin.Text

 'Groupe de securité
 g_iNoGroupe = rstEmploye.Fields("groupe")

 'Nom de l'employé
 g_sEmploye = rstEmploye.Fields("employe")

  Call InitialiserVariablesGroupe

 'Initiale de l'employé
  g_sInitiale = rstEmploye.Fields("initiale")

  Call rstEmploye.Close
  Set rstEmploye = Nothing

  Screen.MousePointer = vbHourglass

 'Fermeture du login
  Call Unload(Me)

 'Ouverture de Dispatch
  Call OuvrirForm(FrmDispatch, False)
  Else
 Call MsgBox("Mot de passe invalide!")
1 End If
Else
 Call MsgBox("L'utilisateur n'existe pas!")
End If
 
 
 Exit Sub

Oups:

 wOups "frmLoginPrincipal", "cmdOK_Click", Err, Err.number, Err.Description

End Sub

Private Sub Form_Activate()

 On Error GoTo Oups

 If txtlogin.Text = "" Then
 Call txtlogin.SetFocus
 Else
 Call txtpasswd.SetFocus
 End If

 Exit Sub

Oups:

 wOups "frmLoginPrincipal", "Form_Activate", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups

 Dim rstConfig As ADODB.Recordset
 Dim sVersion As String
 Dim sDerniereVersion As String

 If OuvrirConnection = True Then
 sVersion = App.Major & "." & Right$("0" & App.Minor, 2) & "." & Right$("0" & App.Revision, 4)

 Set rstConfig = New ADODB.Recordset

 Call rstConfig.Open("SELECT DerniereVersion FROM GrbConfig", g_connData, adOpenDynamic, adLockOptimistic)

 If Not IsNull(rstConfig.Fields("DerniereVersion")) Then
 If rstConfig.Fields("DerniereVersion") <> "" Then
 sDerniereVersion = rstConfig.Fields("DerniereVersion")
  Else
  sDerniereVersion = ""
  End If
  Else
  sDerniereVersion = ""
  End If

  Call rstConfig.Close
  Set rstConfig = Nothing

If sDerniereVersion <> sVersion Then
Call MsgBox("Votre version n'est pas à jour, elle sera installée!", vbOKOnly)

 Call ShellExecute(Me.hwnd, vbNullString, App.Path & "\InstallGRB.exe", vbNullString, "C:\", SW_SHOWNORMAL)

 Call FermerConnection

 End
 End If

 txtlogin.Text = GetSetting("GRB", "Config", "LoginPrincipal", "")
End If

Exit Sub

Oups:

wOups "frmLoginPrincipal", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub txtlogin_GotFocus()

 On Error GoTo Oups

 If txtlogin.Text <> "" Then
 txtlogin.SelStart = 0
 txtlogin.SelLength = Len(txtlogin.Text)
 End If

 Exit Sub

Oups:

 wOups "frmLoginPrincipal", "txtlogin_GotFocus", Err, Err.number, Err.Description
End Sub
