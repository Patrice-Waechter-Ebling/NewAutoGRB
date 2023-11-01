VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAlarme 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alarme"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   ControlBox      =   0   'False
   Icon            =   "frmAlarme.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4215
   Begin VB.CommandButton cmdOK 
      Caption         =   "Valider"
      Height          =   375
      Left            =   3120
      TabIndex        =   9
      Top             =   2520
      Width           =   975
   End
   Begin MSComCtl2.MonthView mvwDate 
      Height          =   2370
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   0
      Appearance      =   1
      StartOfWeek     =   152633345
      CurrentDate     =   38097
   End
   Begin VB.CommandButton cmdDate 
      Caption         =   "..."
      Height          =   255
      Left            =   3600
      TabIndex        =   6
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CheckBox chkRemind 
      BackColor       =   &H00000000&
      Caption         =   "Me le rappeler le :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1575
   End
   Begin MSMask.MaskEdBox mskHeure 
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   2040
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtMessage 
      Height          =   975
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Heure :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Date :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Vous avez une alarme!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmAlarme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_lIDAlarme As Long

Private Sub chkRemind_Click()
 'Active ou désactive les champs
 
 On Error GoTo Oups

 Call ActiverChamps


 Exit Sub

Oups:

 wOups "frmAlarme", "chkRemind_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdOK_Click()
 
 On Error GoTo Oups
 
 Dim rstAlarme As ADODB.Recordset
 Dim sType As String

 Set rstAlarme = New ADODB.Recordset

 Call rstAlarme.Open("SELECT * FROM GrbAlarmes WHERE IDAlarme = " & m_lIDAlarme, g_connData, adOpenDynamic, adLockOptimistic)

 sType = rstAlarme.Fields("TypeCédule")

 If chkRemind.Value = vbChecked Then
 If txtDate.Text <> "" Then
 If mskHeure.Text <> "" Then
 rstAlarme.Fields("Date") = txtDate.Text
 rstAlarme.Fields("Heure") = mskHeure.Text
  rstAlarme.Fields("JourSemaine") = Weekday(txtDate.Text)

  Call rstAlarme.Update
  Else
  Call MsgBox("L'heure est invalide!", vbOKOnly, "Erreur")
  End If
  Else
  Call MsgBox("La date est invalide!", vbOKOnly, "Erreur")
  End If
10 Else
1 Call rstAlarme.Delete

 Call rstAlarme.Update
End If

Call rstAlarme.Close
Set rstAlarme = Nothing

If g_bCeduleOuverte = True Then
 Call frmCédule.RemplirListerJour
 Call frmCédule.RemplirListerSemaine
End If

Call Unload(Me)

Exit Sub

Oups:

1  wOups "frmAlarme", "cmdOK_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Click()

 On Error GoTo Oups

 mvwDate.Visible = False

 Exit Sub

Oups:

 wOups "frmAlarme", "Form_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups

 Call ActiverChamps

 Exit Sub

Oups:

 wOups "frmAlarme", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub ActiverChamps()
 'Active ou désactive les champs
 
 On Error GoTo Oups

 Dim bActif As Boolean

 If chkRemind.Value = vbChecked Then
 bActif = True
 Else
 bActif = False
 End If

 txtDate.Enabled = bActif
 mskHeure.Enabled = bActif
 cmdDate.Enabled = bActif

 Exit Sub

Oups:

  wOups "frmAlarme", "ActiverChamps", Err, Err.number, Err.Description
End Sub

Private Sub mvwDate_DateClick(ByVal DateClicked As Date)

 On Error GoTo Oups

 txtDate.Text = ConvertDate(DateClicked)

 mvwDate.Visible = False

 Exit Sub

Oups:

 wOups "frmAlarme", "mvwDate_DateClick", Err, Err.number, Err.Description
End Sub

Private Sub mvwDate_LostFocus()

 On Error GoTo Oups

 mvwDate.Visible = False

 Exit Sub

Oups:

 wOups "frmAlarme", "mvwDate_LostFocus", Err, Err.number, Err.Description
End Sub

Public Sub Afficher(ByVal lIDAlarme As Long)
 'Affichage de l'alarme
 
 On Error GoTo Oups

 Dim rstAlarme As ADODB.Recordset

 Set rstAlarme = New ADODB.Recordset
 
 m_lIDAlarme = lIDAlarme
 
 'Ouverture de la table
 Call rstAlarme.Open("SELECT * FROM GrbAlarmes WHERE IDAlarme = " & lIDAlarme, g_connData, adOpenDynamic, adLockOptimistic)

 txtMessage.Text = rstAlarme.Fields("Message")

 Call rstAlarme.Close
 Set rstAlarme = Nothing

 Call Me.Show(vbModal)

 Exit Sub

Oups:

 wOups "frmAlarme", "Afficher", Err, Err.number, Err.Description
End Sub

Private Sub mskHeure_GotFocus()

 On Error GoTo Oups
 
 'Format d'heure
 mskHeure.mask = "##:##"

 Exit Sub

Oups:

 wOups "frmAlarme", "mskHeure_GotFocus", Err, Err.number, Err.Description
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

 wOups "frmAlarme", "mskHeure_LostFocus", Err, Err.number, Err.Description
End Sub

Private Sub cmdDate_Click()

 On Error GoTo Oups
 
 'Ouverture du calendrier
 
 'S'il y a une date, la date par défaut est celle-ci, sinon c'est la date
 'd'aujourd'hui
 If Trim$(txtDate.Text) <> vbNullString Then
 mvwDate.Value = txtDate.Text
 Else
 mvwDate.Value = Date
 End If
 
 mvwDate.Visible = True
 
 Call mvwDate.SetFocus

 Exit Sub

Oups:

 wOups "frmAlarme", "cmdDate_Click", Err, Err.number, Err.Description
End Sub
