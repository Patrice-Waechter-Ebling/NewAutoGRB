VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCommentairesProjSoum 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Commentaires"
   ClientHeight    =   10140
   ClientLeft      =   3570
   ClientTop       =   3240
   ClientWidth     =   9930
   Icon            =   "frmCommentairesProjSoum.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10140
   ScaleWidth      =   9930
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdFermer 
      Caption         =   "Fermer"
      Height          =   495
      Left            =   8280
      TabIndex        =   10
      Top             =   9480
      Width           =   1575
   End
   Begin VB.CommandButton cmdCopier 
      Caption         =   "Copier en bas"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton cmdSupprimerTout 
      Caption         =   "Supprimer tout"
      Height          =   495
      Left            =   8280
      TabIndex        =   7
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton cmdVider 
      Caption         =   "Vider"
      Height          =   495
      Left            =   6600
      TabIndex        =   4
      Top             =   8760
      Width           =   1575
   End
   Begin VB.CommandButton cmdAjouter 
      Caption         =   "Ajouter"
      Height          =   495
      Left            =   8280
      TabIndex        =   2
      Top             =   8760
      Width           =   1575
   End
   Begin VB.TextBox txtAjout 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   6000
      Width           =   9735
   End
   Begin MSComctlLib.TreeView tvwCommentaire 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   7435
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Les lignes commencées par ""--"" seront des sous-sections"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   9360
      Width           =   3375
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
      Height          =   255
      Left            =   5880
      TabIndex        =   6
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label lblTitreNoProjSoum 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "# Projet :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4080
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Les lignes commencées par ""-"" seront des sections"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   8760
      Width           =   3375
   End
End
Attribute VB_Name = "frmCommentairesProjSoum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Afficher(ByVal sNoProjSoum As String, ByVal bProjet As Boolean)

 On Error GoTo Oups
 
 If bProjet = True Then
 lblTitreNoProjSoum.Caption = "# Projet : "
 Else
 lblTitreNoProjSoum.Caption = "# Soumission : "
 End If
 
 lblNoProjSoum.Caption = sNoProjSoum

 Call RemplirTreeView

 Call Me.Show(vbModal)

 Exit Sub

Oups:

 wOups "frmCommentairesProjSoum", "Afficher", Err, Err.number, Err.Description
End Sub

Private Sub Cmdajouter_Click()
 
 On Error GoTo Oups

 Dim rstCommentaire As ADODB.Recordset
 Dim arr_sLigne() As String
 Dim sLigne As String
 Dim iKeySection As Integer
 Dim iKeySousSection As Integer
 Dim iCompteur As Integer
 Dim bSousSection As Boolean
 
 If Trim$(txtAjout.Text) <> "" Then
 If tvwCommentaire.Nodes.count = 0 Then
 If Left$(Trim$(txtAjout.Text), 1) = "-" And Left$(Trim$(txtAjout.Text), 2) <> "--" Then
  Set rstCommentaire = New ADODB.Recordset
 
  Call rstCommentaire.Open("SELECT * FROM GrbCommentaires WHERE NoProjSoum = '" & lblNoProjSoum.Caption & "'", g_connData, adOpenDynamic, adLockOptimistic)
 
  arr_sLigne = Split(txtAjout.Text, vbCrLf)

  For iCompteur = 0 To UBound(arr_sLigne)
  sLigne = Trim$(arr_sLigne(iCompteur))

  If sLigne <> "" Then
  Call rstCommentaire.AddNew

  rstCommentaire.Fields("NoProjSoum") = lblNoProjSoum.Caption
 rstCommentaire.Fields("Index") = iCompteur
 
 If Left$(sLigne, 2) = "--" Then
 bSousSection = True

 rstCommentaire.Fields("SousSection") = True

 rstCommentaire.Fields("Relative") = iKeySection

 iKeySousSection = iKeySousSection + 1

 rstCommentaire.Fields("Key") = iKeySousSection

 rstCommentaire.Fields("Commentaire") = Right$(sLigne, Len(sLigne) - 2)
 Else
 If Left$(sLigne, 1) = "-" Then
 rstCommentaire.Fields("Section") = True

 iKeySection = iKeySection + 1

 iKeySousSection = iKeySection * 100

 bSousSection = False

 rstCommentaire.Fields("Key") = iKeySection

 rstCommentaire.Fields("Commentaire") = Right$(sLigne, Len(sLigne) - 1)
 Else
 If bSousSection = True Then
 rstCommentaire.Fields("Relative") = iKeySousSection
1  Else
 rstCommentaire.Fields("Relative") = iKeySection
 End If

 rstCommentaire.Fields("Commentaire") = sLigne
 End If
 End If

 Call rstCommentaire.Update
 End If
 Next

 Call rstCommentaire.Close
 Set rstCommentaire = Nothing

 Call RemplirTreeView
 Else
 Call MsgBox("La première ligne doit absolument être une section!", vbOKOnly, "Erreur")
 End If
Else
 Call MsgBox("Impossible d'ajouter les commentaires, la liste doit être vide!", vbOKOnly, "Erreur")
End If
Else
Call MsgBox("Rien à ajouter!", vbOKOnly, "Erreur")
End If

30 Exit Sub

Oups:

wOups "frmCommentairesProjSoum", "cmdAjouter_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdCopier_Click()

 On Error GoTo Oups

 Dim iCompteur As Integer
 Dim nodComment As Node

 If tvwCommentaire.Nodes.count > 0 Then
 txtAjout.Text = ""

 For iCompteur = 1 To tvwCommentaire.Nodes.count
 Set nodComment = tvwCommentaire.Nodes(iCompteur)

 If nodComment.Bold = True Then
 If txtAjout.Text = "" Then
 txtAjout.Text = "-" & nodComment.Text
 Else
  txtAjout.Text = txtAjout.Text & vbNewLine & "-" & nodComment.Text
  End If
  Else
  If txtAjout.Text = "" Then
  txtAjout.Text = nodComment.Text
  Else
  txtAjout.Text = txtAjout.Text & vbNewLine & nodComment.Text
  End If
 End If
1 Next
End If

Exit Sub

Oups:

wOups "frmCommentairesProjSoum", "cmdCopier_Click", Err, Err.number, Err.Description
End Sub

Private Sub Cmdfermer_Click()

 On Error GoTo Oups

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmCommentairesProjSoum", "cmdFermer_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdSupprimerTout_Click()

 On Error GoTo Oups

 If tvwCommentaire.Nodes.count > 0 Then
 If MsgBox("Voulez-vous vraiment effacer tous les commentaires?", vbYesNo) = vbYes Then
 Call g_connData.Execute("DELETE * FROM GrbCommentaires WHERE NoProjSoum = '" & lblNoProjSoum.Caption & "'")

 Call tvwCommentaire.Nodes.Clear
 End If
 End If

 Exit Sub

Oups:

 wOups "frmCommentairesProjSoum", "cmdSupprimerTout_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdVider_Click()

 On Error GoTo Oups

 txtAjout.Text = ""

 Exit Sub

Oups:

 wOups "frmCommentairesProjSoum", "cmdVider_Click", Err, Err.number, Err.Description
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

wOups "frmCommentairesProjSoum", "RemplirTreeView", Err, Err.number, Err.Description
End Sub
