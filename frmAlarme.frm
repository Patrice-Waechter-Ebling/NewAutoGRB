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
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   0
      Appearance      =   1
      StartOfWeek     =   90243073
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
        
5       On Error GoTo AfficherErreur

10      Call ActiverChamps


15      Exit Sub

AfficherErreur:

20      woups "frmAlarme", "chkRemind_Click", Err, Erl
End Sub

Private Sub cmdOK_Click()
        
5       On Error GoTo AfficherErreur
        
10      Dim rstAlarme As ADODB.Recordset
15      Dim sType     As String

20      Set rstAlarme = New ADODB.Recordset

25      Call rstAlarme.Open("SELECT * FROM GRB_Alarmes WHERE IDAlarme = " & m_lIDAlarme, g_connData, adOpenDynamic, adLockOptimistic)

30      sType = rstAlarme.Fields("TypeCédule")

35      If chkRemind.Value = vbChecked Then
40        If txtDate.Text <> "" Then
45          If mskHeure.Text <> "" Then
50            rstAlarme.Fields("Date") = txtDate.Text
55            rstAlarme.Fields("Heure") = mskHeure.Text
60            rstAlarme.Fields("JourSemaine") = Weekday(txtDate.Text)

65            Call rstAlarme.Update
70          Else
75            Call MsgBox("L'heure est invalide!", vbOKOnly, "Erreur")
80          End If
85        Else
90          Call MsgBox("La date est invalide!", vbOKOnly, "Erreur")
95        End If
100     Else
105       Call rstAlarme.Delete

110       Call rstAlarme.Update
115     End If

120     Call rstAlarme.Close
125     Set rstAlarme = Nothing

130     If g_bCeduleOuverte = True Then
135       Call frmCédule.RemplirListerJour
140       Call frmCédule.RemplirListerSemaine
145     End If

150     Call Unload(Me)

155     Exit Sub

AfficherErreur:

160     woups "frmAlarme", "cmdOK_Click", Err, Erl
End Sub

Private Sub Form_Click()

5       On Error GoTo AfficherErreur

10      mvwDate.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmAlarme", "Form_Click", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

10      Call ActiverChamps

15      Exit Sub

AfficherErreur:

20      woups "frmAlarme", "Form_Load", Err, Erl
End Sub

Private Sub ActiverChamps()
        'Active ou désactive les champs
        
5       On Error GoTo AfficherErreur

10      Dim bActif As Boolean

15      If chkRemind.Value = vbChecked Then
20        bActif = True
25      Else
30        bActif = False
35      End If

40      txtDate.Enabled = bActif
45      mskHeure.Enabled = bActif
50      cmdDate.Enabled = bActif

55      Exit Sub

AfficherErreur:

60      woups "frmAlarme", "ActiverChamps", Err, Erl
End Sub

Private Sub mvwDate_DateClick(ByVal DateClicked As Date)

5       On Error GoTo AfficherErreur

10      txtDate.Text = ConvertDate(DateClicked)

15      mvwDate.Visible = False

20      Exit Sub

AfficherErreur:

25      woups "frmAlarme", "mvwDate_DateClick", Err, Erl
End Sub

Private Sub mvwDate_LostFocus()

5       On Error GoTo AfficherErreur

10      mvwDate.Visible = False

15      Exit Sub

AfficherErreur:

20      woups "frmAlarme", "mvwDate_LostFocus", Err, Erl
End Sub

Public Sub Afficher(ByVal lIDAlarme As Long)
        'Affichage de l'alarme
        
5       On Error GoTo AfficherErreur

10      Dim rstAlarme As ADODB.Recordset

15      Set rstAlarme = New ADODB.Recordset
 
20      m_lIDAlarme = lIDAlarme
 
        'Ouverture de la table
25      Call rstAlarme.Open("SELECT * FROM GRB_Alarmes WHERE IDAlarme = " & lIDAlarme, g_connData, adOpenDynamic, adLockOptimistic)

30      txtMessage.Text = rstAlarme.Fields("Message")

35      Call rstAlarme.Close
40      Set rstAlarme = Nothing

45      Call Me.Show(vbModal)

50      Exit Sub

AfficherErreur:

55      woups "frmAlarme", "Afficher", Err, Erl
End Sub

Private Sub mskHeure_GotFocus()

5       On Error GoTo AfficherErreur
        
        'Format d'heure
10      mskHeure.mask = "##:##"

15      Exit Sub

AfficherErreur:

20      woups "frmAlarme", "mskHeure_GotFocus", Err, Erl
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

35      woups "frmAlarme", "mskHeure_LostFocus", Err, Erl
End Sub

Private Sub cmdDate_Click()

5       On Error GoTo AfficherErreur
        
        'Ouverture du calendrier
  
        'S'il y a une date, la date par défaut est celle-ci, sinon c'est la date
        'd'aujourd'hui
10      If Trim$(txtDate.Text) <> vbNullString Then
15        mvwDate.Value = txtDate.Text
20      Else
25        mvwDate.Value = Date
30      End If
  
35      mvwDate.Visible = True
  
40      Call mvwDate.SetFocus

45      Exit Sub

AfficherErreur:

50      woups "frmAlarme", "cmdDate_Click", Err, Erl
End Sub
