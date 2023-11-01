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
   MinButton       =   0   'False
   Picture         =   "frmChoixQteBoite.frx":0000
   ScaleHeight     =   2790
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
  
5       On Error GoTo AfficherErreur

10      lblQuestion.Caption = "Quelle est la quantité par boîte pour la pièce " & sNoPiece & "?"

15      Call Me.Show(vbModal)

20      Exit Sub

AfficherErreur:
  
25      woups "frmChoixQteBoite", "Afficher", Err, Erl
End Sub

Private Sub chkQteBoite_Click()

5       On Error GoTo AfficherErreur

10      If chkQteBoite.Value = vbChecked Then
15        txtQteBoite.Enabled = True
20      Else
25        txtQteBoite.Text = ""
    
30        txtQteBoite.Enabled = False
35      End If

40      Exit Sub

AfficherErreur:

45      woups "frmChoixQteBoite", "chkQteBoite_Click", Err, Erl
End Sub

Private Sub cmdOK_Click()
  
5       On Error GoTo AfficherErreur

10      g_sQteBoite = txtQteBoite.Text

15      If chkQteBoite.Value = vbChecked Then
20        g_bQteBoite = True
25      Else
30        g_bQteBoite = False
35      End If

40      Call Unload(Me)

45      Exit Sub

AfficherErreur:
  
50      woups "frmChoixQteBoite", "cmdOK_Click", Err, Erl
End Sub

Private Sub txtQteBoite_LostFocus()

5       On Error GoTo AfficherErreur

10      If chkQteBoite.Value = vbChecked Then
15        txtQteBoite.Text = Replace(txtQteBoite.Text, ".", ",")

20        If Not IsNumeric(txtQteBoite.Text) Or txtQteBoite.Text = "0" Then
25          txtQteBoite.Text = "0"
30        End If
35      End If

40      Exit Sub

AfficherErreur:

45      woups "frmChoixQteBoite", "txtQteBoite_LostFocus", Err, Erl
End Sub
