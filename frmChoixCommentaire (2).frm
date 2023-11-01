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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   9930
   ShowInTaskbar   =   0   'False
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

 On Error GoTo Oups
 
 lblNoProjSoum.Caption = sNoProjSoum

 Call RemplirTreeView

 Call Me.Show(vbModal)

 Exit Sub

Oups:

 wOups "frmChoixCommentaire", "Afficher", Err, Err.number, Err.Description
End Sub

Private Sub RemplirTreeView()

 On Error GoTo Oups

 Dim rstCommentaire As ADODB.Recordset
 Dim itmCommentaire As Node

 Call tvwCommentaire.Nodes.Clear

 Set rstCommentaire = New ADODB.Recordset

 Call rstCommentaire.Open("SELECT * FROM GrbCommentaires WHERE NoProjSoum = '" & lblNoProjSoum.Caption & "'", g_connData, adOpenDynamic, adLockOptimistic)

 Do While Not rstCommentaire.EOF
 If Not IsNull(rstCommentaire.Fields("Commentaire")) Then
 If rstCommentaire.Fields("Section") = True Then
 Set itmCommentaire = tvwCommentaire.Nodes.Add(, , "KEY" & rstCommentaire.Fields("Key"), rstCommentaire.Fields("Commentaire"))

 itmCommentaire.Bold = True
  Else
  If rstCommentaire.Fields("SousSection") = True Then
  Set itmCommentaire = tvwCommentaire.Nodes.Add("KEY" & rstCommentaire.Fields("Relative"), tvwChild, "KEY" & rstCommentaire.Fields("Key"), rstCommentaire.Fields("Commentaire"))

  itmCommentaire.Bold = True
  Else
  Set itmCommentaire = tvwCommentaire.Nodes.Add("KEY" & rstCommentaire.Fields("Relative"), tvwChild, , rstCommentaire.Fields("Commentaire"))
  End If
  End If

 itmCommentaire.Tag = rstCommentaire.Fields("ID")
1 End If

 Call rstCommentaire.MoveNext
Loop

Call rstCommentaire.Close
Set rstCommentaire = Nothing

Exit Sub

Oups:

wOups "frmChoixCommentaire", "RemplirTreeView", Err, Err.number, Err.Description
End Sub

Private Sub cmdOK_Click()

 On Error GoTo Oups

 Dim sCommentaire As String
 Dim iCompteur As Integer
 Dim sSection As String
 Dim nodComment As Node

 If tvwCommentaire.Nodes.count > 0 Then
 For iCompteur = 1 To tvwCommentaire.Nodes.count
 Set nodComment = tvwCommentaire.Nodes(iCompteur)

 If nodComment.Bold = True Then
 If CInt(Replace(nodComment.Key, "KEY", "")) > 100 Then
 sSection = sSection & " / " & nodComment.Text
  Else
  sSection = nodComment.Text
  End If
  Else
  If nodComment.Checked = True Then
  If sCommentaire = "" Then
  sCommentaire = sSection & " / " & nodComment.Text
  Else
 sCommentaire = sCommentaire & vbNewLine & sSection & " / " & nodComment.Text
 End If
 End If
 End If
 Next

 frmPunch.sCommentaire = sCommentaire
End If

Call Unload(Me)

Exit Sub

Oups:

wOups "frmChoixCommentaire", "cmdOK_Click", Err, Err.number, Err.Description
End Sub
