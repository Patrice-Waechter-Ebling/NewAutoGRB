VERSION 5.00
Begin VB.Form frmChoixLocalisation 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbLocalisation 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   3615
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblQuestion 
      BackStyle       =   0  'Transparent
      Caption         =   "Dans quelle localisation ?"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3615
   End
End
Attribute VB_Name = "frmChoixLocalisation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_eCatalogue As enumCatalogue

Public Sub Afficher(ByVal eCatalogue As enumCatalogue, ByVal sNoPiece As String)
 
 On Error GoTo Oups

 m_eCatalogue = eCatalogue

 lblQuestion.Caption = "Quelle est la localisation de la pièce " & sNoPiece & "?"

 Call RemplirComboLocalisation

 Call Me.Show(vbModal)

 Exit Sub

Oups:
 
 wOups "frmChoixLocalisation", "Afficher", Err, Err.number, Err.Description
End Sub

Private Sub cmdOK_Click()

 On Error GoTo Oups

 g_sLocalisation = cmbLocalisation.Text

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmChoixLocalisation", "cmdOK_Click", Err, Err.number, Err.Description
End Sub

Private Sub RemplirComboLocalisation()

 On Error GoTo Oups

 'Remplir le combo des localisations
 Dim rstInv As ADODB.Recordset
 
 'Il faut vider le combo avant de le remplir
 Call cmbLocalisation.Clear

 Call cmbLocalisation.AddItem("")

 Set rstInv = New ADODB.Recordset
 
 If m_eCatalogue = ELECTRIQUE Then
 Call rstInv.Open("SELECT DISTINCT Localisation FROM GrbInventaireElec", g_connData, adOpenDynamic, adLockOptimistic)
 Else
 Call rstInv.Open("SELECT DISTINCT Localisation FROM GrbInventaireMec", g_connData, adOpenDynamic, adLockOptimistic)
 End If

 'Tant que ce n'est pas la fin des enregistrements
 Do While Not rstInv.EOF
  If Trim(rstInv.Fields("Localisation")) <> "" Then
  Call cmbLocalisation.AddItem(rstInv.Fields("Localisation"))
  End If
 
  Call rstInv.MoveNext
  Loop
 
  Call rstInv.Close
  Set rstInv = Nothing
 
  Exit Sub

Oups:

10 wOups "frmChoixLocalisation", "RemplirComboContactFRS", Err, Err.number, Err.Description
End Sub

