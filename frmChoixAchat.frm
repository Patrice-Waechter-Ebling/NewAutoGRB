VERSION 5.00
Begin VB.Form frmChoixAchat 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Achat"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3255
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmChoixAchat.frx":0000
   ScaleHeight     =   3105
   ScaleWidth      =   3255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdElectrique 
      Caption         =   "Électrique"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton cmdMecanique 
      Caption         =   "Mécanique"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdFermer 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   2640
      Width           =   1575
   End
End
Attribute VB_Name = "frmChoixAchat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdElectrique_Click()

10      On Error GoTo AfficherErreur

20      Call frmAchat.Afficher(ELECTRIQUE)

30      Exit Sub

AfficherErreur:

40      Call AfficherErreur(Me, "cmdElectrique_Click", Err, Erl)
End Sub

Private Sub cmdFermer_Click()

10      On Error GoTo AfficherErreur

20      Call Unload(Me)

30      Exit Sub

AfficherErreur:

40      Call AfficherErreur(Me, "cmdFermer_Click", Err, Erl)
End Sub

Private Sub cmdMecanique_Click()

10      On Error GoTo AfficherErreur

20      Call frmAchat.Afficher(MECANIQUE)

30      Exit Sub

AfficherErreur:

40      Call AfficherErreur(Me, "cmdMecanique_Click", Err, Erl)
End Sub
