VERSION 5.00
Begin VB.Form DialogPassMaster 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MOT DE PASSE REQUIS!!"
   ClientHeight    =   1140
   ClientLeft      =   3870
   ClientTop       =   4050
   ClientWidth     =   4995
   ControlBox      =   0   'False
   Icon            =   "DialogPassMaster.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Veuillez entrer votre mot de passe :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "DialogPassMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelButton_Click()
       
5       On Error GoTo VerifierErreur

        'Fermeture de la fenêtre
10      Call Unload(Me)
  
VerifierErreur:
  
15      Call AfficherErreur(Me, "CancelButton_Click", Err, Erl)
End Sub

Private Sub OKButton_Click()

5       On Error GoTo AfficherErreur
 
        'Vérification du mot de passe
10      If UCase(txtPassword.Text) = UCase(sCfgPass) Then
15        Call Unload(Me)
    
          'Ouverture du chemin de la BD
20        Call OuvrirForm(frmCheminBD, True)
25      End If

30      Exit Sub

AfficherErreur:

35     Call AfficherErreur(Me, "OKButton_Click", Err, Erl)
End Sub
