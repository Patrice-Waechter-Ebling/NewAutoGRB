VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCommentairesProjSoum 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Commentaires"
   ClientHeight    =   10140
   ClientLeft      =   3570
   ClientTop       =   3240
   ClientWidth     =   9930
   Icon            =   "frmCommentairesProjSoum.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCommentairesProjSoum.frx":000C
   ScaleHeight     =   10140
   ScaleWidth      =   9930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      ForeColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5880
      TabIndex        =   6
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label lblTitreNoProjSoum 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00FFFFFF&
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

5       On Error GoTo AfficherErreur
  
10      If bProjet = True Then
15        lblTitreNoProjSoum.Caption = "# Projet : "
20      Else
25        lblTitreNoProjSoum.Caption = "# Soumission : "
30      End If
  
35      lblNoProjSoum.Caption = sNoProjSoum

40      Call RemplirTreeView

45      Call Me.Show(vbModal)

50      Exit Sub

AfficherErreur:

55      woups "frmCommentairesProjSoum", "Afficher", Err, Erl
End Sub

Private Sub Cmdajouter_Click()
  
5       On Error GoTo AfficherErreur

10      Dim rstCommentaire  As ADODB.Recordset
15      Dim arr_sLigne()    As String
20      Dim sLigne          As String
25      Dim iKeySection     As Integer
30      Dim iKeySousSection As Integer
35      Dim iCompteur       As Integer
40      Dim bSousSection    As Boolean
  
45      If Trim$(txtAjout.Text) <> "" Then
50        If tvwCommentaire.Nodes.count = 0 Then
55          If Left$(Trim$(txtAjout.Text), 1) = "-" And Left$(Trim$(txtAjout.Text), 2) <> "--" Then
60            Set rstCommentaire = New ADODB.Recordset
  
65            Call rstCommentaire.Open("SELECT * FROM GRB_Commentaires WHERE NoProjSoum = '" & lblNoProjSoum.Caption & "'", g_connData, adOpenDynamic, adLockOptimistic)
        
70            arr_sLigne = Split(txtAjout.Text, vbCrLf)

75            For iCompteur = 0 To UBound(arr_sLigne)
80              sLigne = Trim$(arr_sLigne(iCompteur))

85              If sLigne <> "" Then
90                Call rstCommentaire.AddNew

95                rstCommentaire.Fields("NoProjSoum") = lblNoProjSoum.Caption
100               rstCommentaire.Fields("Index") = iCompteur
 
105               If Left$(sLigne, 2) = "--" Then
110                 bSousSection = True

115                 rstCommentaire.Fields("SousSection") = True

120                 rstCommentaire.Fields("Relative") = iKeySection

125                 iKeySousSection = iKeySousSection + 1

130                 rstCommentaire.Fields("Key") = iKeySousSection

135                 rstCommentaire.Fields("Commentaire") = Right$(sLigne, Len(sLigne) - 2)
140               Else
145                 If Left$(sLigne, 1) = "-" Then
150                   rstCommentaire.Fields("Section") = True

155                   iKeySection = iKeySection + 1

160                   iKeySousSection = iKeySection * 100

165                   bSousSection = False

170                   rstCommentaire.Fields("Key") = iKeySection

175                   rstCommentaire.Fields("Commentaire") = Right$(sLigne, Len(sLigne) - 1)
180                 Else
185                   If bSousSection = True Then
190                     rstCommentaire.Fields("Relative") = iKeySousSection
195                   Else
200                     rstCommentaire.Fields("Relative") = iKeySection
205                   End If

210                   rstCommentaire.Fields("Commentaire") = sLigne
215                 End If
220               End If

225               Call rstCommentaire.Update
230             End If
235           Next

240           Call rstCommentaire.Close
245           Set rstCommentaire = Nothing

250           Call RemplirTreeView
255         Else
260           Call MsgBox("La première ligne doit absolument être une section!", vbOKOnly, "Erreur")
265         End If
270       Else
275         Call MsgBox("Impossible d'ajouter les commentaires, la liste doit être vide!", vbOKOnly, "Erreur")
280       End If
285     Else
290       Call MsgBox("Rien à ajouter!", vbOKOnly, "Erreur")
295     End If

300     Exit Sub

AfficherErreur:

305     woups "frmCommentairesProjSoum", "cmdAjouter_Click", Err, Erl
End Sub

Private Sub cmdCopier_Click()

5       On Error GoTo AfficherErreur

10      Dim iCompteur  As Integer
15      Dim nodComment As Node

20      If tvwCommentaire.Nodes.count > 0 Then
25        txtAjout.Text = ""

30        For iCompteur = 1 To tvwCommentaire.Nodes.count
35          Set nodComment = tvwCommentaire.Nodes(iCompteur)

40          If nodComment.Bold = True Then
45            If txtAjout.Text = "" Then
50              txtAjout.Text = "-" & nodComment.Text
55            Else
60              txtAjout.Text = txtAjout.Text & vbNewLine & "-" & nodComment.Text
65            End If
70          Else
75            If txtAjout.Text = "" Then
80              txtAjout.Text = nodComment.Text
85            Else
90              txtAjout.Text = txtAjout.Text & vbNewLine & nodComment.Text
95            End If
100         End If
105       Next
110     End If

115     Exit Sub

AfficherErreur:

120     woups "frmCommentairesProjSoum", "cmdCopier_Click", Err, Erl
End Sub

Private Sub Cmdfermer_Click()

5       On Error GoTo AfficherErreur

10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmCommentairesProjSoum", "cmdFermer_Click", Err, Erl
End Sub

Private Sub cmdSupprimerTout_Click()

5       On Error GoTo AfficherErreur

10      If tvwCommentaire.Nodes.count > 0 Then
15        If MsgBox("Voulez-vous vraiment effacer tous les commentaires?", vbYesNo) = vbYes Then
20          Call g_connData.Execute("DELETE * FROM GRB_Commentaires WHERE NoProjSoum = '" & lblNoProjSoum.Caption & "'")

25          Call tvwCommentaire.Nodes.Clear
30        End If
35      End If

40      Exit Sub

AfficherErreur:

45      woups "frmCommentairesProjSoum", "cmdSupprimerTout_Click", Err, Erl
End Sub

Private Sub cmdVider_Click()

5       On Error GoTo AfficherErreur

10      txtAjout.Text = ""

15      Exit Sub

AfficherErreur:

20      woups "frmCommentairesProjSoum", "cmdVider_Click", Err, Erl
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

100          itmCommentaire.Tag = rstCommentaire.Fields("ID")
105       End If

110       Call rstCommentaire.MoveNext
115     Loop

120     Call rstCommentaire.Close
125     Set rstCommentaire = Nothing

130     Exit Sub

AfficherErreur:

135     woups "frmCommentairesProjSoum", "RemplirTreeView", Err, Erl
End Sub
