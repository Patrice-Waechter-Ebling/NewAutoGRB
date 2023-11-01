VERSION 5.00
Begin VB.Form frmChoixQteBoite 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   2790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtQteBoite 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.CheckBox chkQteBoite 
      BackColor       =   &H00000000&
      Caption         =   "Commande par boîte"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label lblQuestion 
      BackStyle       =   0  'Transparent
      Caption         =   "Quelle est la quantité par boîte pour la pièce ?"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label lblQteBoite 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantité :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   1800
      Width           =   855
   End
End
Attribute VB_Name = "frmChoixQteBoite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Afficher(ByVal sNoPiece As String)
 
 On Error GoTo Oups

 lblQuestion.Caption = "Quelle est la quantité par boîte pour la pièce " & sNoPiece & "?"

 Call Me.Show(vbModal)

 Exit Sub

Oups:
 
 wOups "frmChoixQteBoite", "Afficher", Err, Err.number, Err.Description
End Sub

Private Sub chkQteBoite_Click()

 On Error GoTo Oups

 If chkQteBoite.Value = vbChecked Then
 txtQteBoite.Enabled = True
 Else
 txtQteBoite.Text = ""
 
 txtQteBoite.Enabled = False
 End If

 Exit Sub

Oups:

 wOups "frmChoixQteBoite", "chkQteBoite_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdOK_Click()
 
 On Error GoTo Oups

 g_sQteBoite = txtQteBoite.Text

 If chkQteBoite.Value = vbChecked Then
 g_bQteBoite = True
 Else
 g_bQteBoite = False
 End If

 Call Unload(Me)

 Exit Sub

Oups:
 
 wOups "frmChoixQteBoite", "cmdOK_Click", Err, Err.number, Err.Description
End Sub

Private Sub txtQteBoite_LostFocus()

 On Error GoTo Oups

 If chkQteBoite.Value = vbChecked Then
 txtQteBoite.Text = Replace(txtQteBoite.Text, ".", ",")

 If Not IsNumeric(txtQteBoite.Text) Or txtQteBoite.Text = "0" Then
 txtQteBoite.Text = "0"
 End If
 End If

 Exit Sub

Oups:

 wOups "frmChoixQteBoite", "txtQteBoite_LostFocus", Err, Err.number, Err.Description
End Sub
