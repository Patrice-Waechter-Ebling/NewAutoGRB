VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChoixCommentaire 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Commentaires"
   ClientHeight    =   9390
   ClientLeft      =   3570
   ClientTop       =   3240
   ClientWidth     =   9930
   ControlBox      =   0   'False
   Icon            =   "frmChoixCommentaire.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmChoixCommentaire.frx":000C
   ScaleHeight     =   9390
   ScaleWidth      =   9930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   8280
      TabIndex        =   1
      Top             =   8760
      Width           =   1575
   End
   Begin MSComctlLib.TreeView tvwCommentaire 
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   13361
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
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
   End
   Begin VB.Label lblNoProjSoum 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5880
      TabIndex        =   3
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label lblTitreNoProjSoum 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Numéro :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4080
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmChoixCommentaire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Afficher(ByVal sNoProjSoum As String)

5       On Error GoTo AfficherErreur
    
10      lblNoProjSoum.Caption = sNoProjSoum

15      Call RemplirTreeView

20      Call Me.Show(vbModal)

25      Exit Sub

AfficherErreur:

30      woups "frmChoixCommentaire", "Afficher", Err, Erl
End Sub

Private Sub RemplirTreeView()

5       On Error GoTo AfficherErreur

10      Dim rstCommentaire As ADODB.Recordset
15      Dim itmCommentaire As Node

20      Call tvwCommentaire.Nodes.Clear

25      Set rstCommentaire = New ADODB.Recordset

30      Call rstCommentaire.Open("SELECT * FROM GRB_Commentaires WHERE NoProjSoum = '" & lblNoProjSoum.Caption & "'", g_connData, adOpenDynamic, adLockOptimistic)

35      Do While Not rstCommentaire.EOF
40        If Not IsNull(rstCommentaire.Fields("Commentaire")) Then
45          If rstCommentaire.Fields("Section") = True Then
50            Set itmCommentaire = tvwCommentaire.Nodes.Add(, , "KEY" & rstCommentaire.Fields("Key"), rstCommentaire.Fields("Commentaire"))

55            itmCommentaire.Bold = True
60          Else
65            If rstCommentaire.Fields("SousSection") = True Then
70              Set itmCommentaire = tvwCommentaire.Nodes.Add("KEY" & rstCommentaire.Fields("Relative"), tvwChild, "KEY" & rstCommentaire.Fields("Key"), rstCommentaire.Fields("Commentaire"))

75              itmCommentaire.Bold = True
80            Else
85              Set itmCommentaire = tvwCommentaire.Nodes.Add("KEY" & rstCommentaire.Fields("Relative"), tvwChild, , rstCommentaire.Fields("Commentaire"))
90            End If
95          End If

100         itmCommentaire.Tag = rstCommentaire.Fields("ID")
105       End If

110       Call rstCommentaire.MoveNext
115     Loop

120     Call rstCommentaire.Close
125     Set rstCommentaire = Nothing

130     Exit Sub

AfficherErreur:

135     woups "frmChoixCommentaire", "RemplirTreeView", Err, Erl
End Sub

Private Sub cmdOK_Click()

5       On Error GoTo AfficherErreur

10      Dim sCommentaire As String
15      Dim iCompteur    As Integer
20      Dim sSection     As String
25      Dim nodComment   As Node

30      If tvwCommentaire.Nodes.count > 0 Then
35        For iCompteur = 1 To tvwCommentaire.Nodes.count
40          Set nodComment = tvwCommentaire.Nodes(iCompteur)

45          If nodComment.Bold = True Then
50            If CInt(Replace(nodComment.Key, "KEY", "")) > 100 Then
55              sSection = sSection & " / " & nodComment.Text
60            Else
65              sSection = nodComment.Text
70            End If
75          Else
80            If nodComment.Checked = True Then
85              If sCommentaire = "" Then
90                sCommentaire = sSection & " / " & nodComment.Text
95              Else
100               sCommentaire = sCommentaire & vbNewLine & sSection & " / " & nodComment.Text
105             End If
110           End If
115         End If
120       Next

125       frmPunch.sCommentaire = sCommentaire
130     End If

135     Call Unload(Me)

140     Exit Sub

AfficherErreur:

145     woups "frmChoixCommentaire", "cmdOK_Click", Err, Erl
End Sub
