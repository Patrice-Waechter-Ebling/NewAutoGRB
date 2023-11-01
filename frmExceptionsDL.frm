VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExceptionsDL 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exceptions des listes de distribution"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmExceptionsDL.frx":0000
   ScaleHeight     =   5745
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSupprimer 
      Caption         =   "Supprimer"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdAjouter 
      Caption         =   "Ajouter"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   1440
      Width           =   1455
   End
   Begin MSComctlLib.ListView lvwExceptions 
      Height          =   4215
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   7435
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Courriel"
         Object.Width           =   6703
      EndProperty
   End
End
Attribute VB_Name = "frmExceptionsDL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdajouter_Click()

5       On Error GoTo AfficherErreur

10      Dim sAdresse      As String
15      Dim rstExceptions As ADODB.Recordset

20      sAdresse = InputBox("Quel est l'adresse à ajouter ?")

25      If StrPtr(sAdresse) <> 0 Then
30        If ValiderAdresse(sAdresse) = True Then
35          Set rstExceptions = New ADODB.Recordset

40          Call rstExceptions.Open("SELECT * FROM GRB_ExceptionsDL WHERE [Exception] = '" & Replace(sAdresse, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)

45          If rstExceptions.EOF Then
50            Call rstExceptions.AddNew

55            rstExceptions.Fields("Exception") = sAdresse

60            Call rstExceptions.Update

65            Call RemplirListBoxExceptions
70          Else
75            Call MsgBox("Ce courriel est déjà dans la liste!", vbOKOnly, "Erreur")
80          End If

85          Call rstExceptions.Close
90          Set rstExceptions = Nothing
95        Else
100         Call MsgBox("Adresse invalide!", vbOKOnly, "Erreur")
105       End If
110     End If

115     Exit Sub

AfficherErreur:

120     woups "frmExceptionsDL", "cmdAjouter_Click", Err, Erl
End Sub

Private Function ValiderAdresse(ByVal sAdresse As String) As Boolean

5       On Error GoTo AfficherErreur

10      Dim bValide As Boolean

15      bValide = True

20      If Len(sAdresse) < 5 Then
25        bValide = False
30      End If

35      If bValide = True Then
40        If InStr(1, sAdresse, "@") = 0 Then
45          bValide = False
50        End If
55      End If

60      If bValide = True Then
65        If InStr(InStr(1, sAdresse, "@") + 1, sAdresse, ".") = 0 Then
70          bValide = False
75        End If
80      End If

85      If bValide = True Then
90        If Left$(sAdresse, 1) = "@" Then
95          bValide = False
100       End If
105     End If

110     If bValide = True Then
115       If Right$(sAdresse, 1) = "." Then
120         bValide = False
125       End If
130     End If

135     ValiderAdresse = bValide


140     Exit Function

AfficherErreur:

145     woups "frmExceptionsDL", "ValiderAdresse", Err, Erl
End Function


Private Sub cmdsupprimer_Click()

5       On Error GoTo AfficherErreur

10      Call SupprimerCourriel

15      Exit Sub

AfficherErreur:

20      woups "frmExceptionsDL", "cmdSupprimer_Click", Err, Erl
End Sub

Private Sub SupprimerCourriel()

5       On Error GoTo AfficherErreur

10      If lvwExceptions.ListItems.count > 0 Then
15        If MsgBox("Voulez-vous vraiment effacer l'adresse " & lvwExceptions.SelectedItem.Text & " ? ", vbYesNo) = vbYes Then
20          Call g_connData.Execute("DELETE * FROM GRB_ExceptionsDL WHERE ID = " & lvwExceptions.SelectedItem.Tag)

25          Call RemplirListBoxExceptions
30        End If
35      End If

40      Exit Sub

AfficherErreur:

45      woups "frmExceptionsDL", "SupprimerCourriel", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

10      Call RemplirListBoxExceptions

15      Exit Sub

AfficherErreur:

20      woups "frmExceptionsDL", "Form_Load", Err, Erl
End Sub

Private Sub RemplirListBoxExceptions()

5       On Error GoTo AfficherErreur

10      Dim rstExceptions As ADODB.Recordset
15      Dim itmException  As ListItem

20      Call lvwExceptions.ListItems.Clear

25      Set rstExceptions = New ADODB.Recordset

30      Call rstExceptions.Open("SELECT * FROM GRB_ExceptionsDL ORDER BY [Exception]", g_connData, adOpenForwardOnly, adLockReadOnly)

35      Do While Not rstExceptions.EOF
40        Set itmException = lvwExceptions.ListItems.Add

45        itmException.Text = rstExceptions.Fields("Exception")
50        itmException.Tag = rstExceptions.Fields("ID")

55        Call rstExceptions.MoveNext
60      Loop

65      Call rstExceptions.Close
70      Set rstExceptions = Nothing

75      Exit Sub

AfficherErreur:

80      woups "frmExceptionsDL", "RemplirListBoxExceptions", Err, Erl
End Sub

Private Sub lvwExceptions_KeyDown(KeyCode As Integer, Shift As Integer)

5       On Error GoTo AfficherErreur

10      If KeyCode = vbKeyDelete Then
15        Call SupprimerCourriel
20      End If

25      Exit Sub

AfficherErreur:

30      woups "frmExceptionsDL", "lvwExceptions_KeyDown", Err, Erl
End Sub
