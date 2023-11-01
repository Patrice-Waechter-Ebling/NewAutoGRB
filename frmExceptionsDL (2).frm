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
   MDIChild        =   -1  'True
   ScaleHeight     =   5745
   ScaleWidth      =   6330
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

 On Error GoTo Oups

 Dim sAdresse As String
 Dim rstExceptions As ADODB.Recordset

 sAdresse = InputBox("Quel est l'adresse à ajouter ?")

 If StrPtr(sAdresse) <> 0 Then
 If ValiderAdresse(sAdresse) = True Then
 Set rstExceptions = New ADODB.Recordset

 Call rstExceptions.Open("SELECT * FROM GrbExceptionsDL WHERE [Exception] = '" & Replace(sAdresse, "'", "''") & "'", g_connData, adOpenDynamic, adLockOptimistic)

 If rstExceptions.EOF Then
 Call rstExceptions.AddNew

 rstExceptions.Fields("Exception") = sAdresse

  Call rstExceptions.Update

  Call RemplirListBoxExceptions
  Else
  Call MsgBox("Ce courriel est déjà dans la liste!", vbOKOnly, "Erreur")
  End If

  Call rstExceptions.Close
  Set rstExceptions = Nothing
  Else
 Call MsgBox("Adresse invalide!", vbOKOnly, "Erreur")
1 End If
End If

Exit Sub

Oups:

wOups "frmExceptionsDL", "cmdAjouter_Click", Err, Err.number, Err.Description
End Sub

Private Function ValiderAdresse(ByVal sAdresse As String) As Boolean

 On Error GoTo Oups

 Dim bValide As Boolean

 bValide = True

 If Len(sAdresse) < 5 Then
 bValide = False
 End If

 If bValide = True Then
 If InStr(1, sAdresse, "@") = 0 Then
 bValide = False
 End If
 End If

  If bValide = True Then
  If InStr(InStr(1, sAdresse, "@") + 1, sAdresse, ".") = 0 Then
  bValide = False
  End If
  End If

  If bValide = True Then
  If Left$(sAdresse, 1) = "@" Then
  bValide = False
End If
End If

If bValide = True Then
 If Right$(sAdresse, 1) = "." Then
 bValide = False
 End If
End If

ValiderAdresse = bValide


Exit Function

Oups:

wOups "frmExceptionsDL", "ValiderAdresse", Err, Err.number, Err.Description
End Function


Private Sub cmdsupprimer_Click()

 On Error GoTo Oups

 Call SupprimerCourriel

 Exit Sub

Oups:

 wOups "frmExceptionsDL", "cmdSupprimer_Click", Err, Err.number, Err.Description
End Sub

Private Sub SupprimerCourriel()

 On Error GoTo Oups

 If lvwExceptions.ListItems.count > 0 Then
 If MsgBox("Voulez-vous vraiment effacer l'adresse " & lvwExceptions.SelectedItem.Text & " ? ", vbYesNo) = vbYes Then
 Call g_connData.Execute("DELETE * FROM GrbExceptionsDL WHERE ID = " & lvwExceptions.SelectedItem.Tag)

 Call RemplirListBoxExceptions
 End If
 End If

 Exit Sub

Oups:

 wOups "frmExceptionsDL", "SupprimerCourriel", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups

 Call RemplirListBoxExceptions

 Exit Sub

Oups:

 wOups "frmExceptionsDL", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub RemplirListBoxExceptions()

 On Error GoTo Oups

 Dim rstExceptions As ADODB.Recordset
 Dim itmException As ListItem

 Call lvwExceptions.ListItems.Clear

 Set rstExceptions = New ADODB.Recordset

 Call rstExceptions.Open("SELECT * FROM GrbExceptionsDL ORDER BY [Exception]", g_connData, adOpenForwardOnly, adLockReadOnly)

 Do While Not rstExceptions.EOF
 Set itmException = lvwExceptions.ListItems.Add

 itmException.Text = rstExceptions.Fields("Exception")
 itmException.Tag = rstExceptions.Fields("ID")

 Call rstExceptions.MoveNext
  Loop

  Call rstExceptions.Close
  Set rstExceptions = Nothing

  Exit Sub

Oups:

  wOups "frmExceptionsDL", "RemplirListBoxExceptions", Err, Err.number, Err.Description
End Sub

Private Sub lvwExceptions_KeyDown(KeyCode As Integer, Shift As Integer)

 On Error GoTo Oups

 If KeyCode = vbKeyDelete Then
 Call SupprimerCourriel
 End If

 Exit Sub

Oups:

 wOups "frmExceptionsDL", "lvwExceptions_KeyDown", Err, Err.number, Err.Description
End Sub
