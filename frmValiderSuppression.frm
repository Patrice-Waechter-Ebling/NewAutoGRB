VERSION 5.00
Begin VB.Form frmValiderSuppression 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4035
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   5910
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtValidation 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdAnnuler 
      Caption         =   "Annuler"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdValider 
      Caption         =   "Valider"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SUPPRIMER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblValidation 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Veuillez réécrire le code de gauche dans la case de droite."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   5655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Voulez-vous vraiment continuer ?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   5655
   End
   Begin VB.Label lblNumero 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "M73000-06"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5655
   End
   Begin VB.Label lblAction 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cette action va"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label lblProjSoum 
      BackStyle       =   0  'Transparent
      Caption         =   "la soumission"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmValiderSuppression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_frmSource As Form

Public Sub Afficher(ByVal bProjet As Boolean, ByVal sNumero As String, ByVal frmSource As Form)
  If bProjet = True Then
    lblProjSoum.Caption = "le projet"
  Else
    lblProjSoum.Caption = "la soumission"
  End If
  
  lblNumero.Caption = sNumero
  
  Set m_frmSource = frmSource
  
  Call AfficherNoValidation
  
  Call Me.Show(vbModal)
End Sub

Private Sub AfficherNoValidation()
  Dim iRandom     As Integer
  Dim iCompteur   As Integer
  Dim sValidation As String

  Call Randomize
  
  For iCompteur = 1 To 3
    iRandom = Int(Rnd * 26) + 1
  
    Select Case iRandom
      Case 1:  sValidation = sValidation & "A"
      Case 2:  sValidation = sValidation & "B"
      Case 3:  sValidation = sValidation & "C"
      Case 4:  sValidation = sValidation & "D"
      Case 5:  sValidation = sValidation & "E"
      Case 6:  sValidation = sValidation & "F"
      Case 7:  sValidation = sValidation & "G"
      Case 8:  sValidation = sValidation & "H"
      Case 9:  sValidation = sValidation & "I"
      Case 10: sValidation = sValidation & "J"
      Case 11: sValidation = sValidation & "K"
      Case 12: sValidation = sValidation & "L"
      Case 13: sValidation = sValidation & "M"
      Case 14: sValidation = sValidation & "N"
      Case 15: sValidation = sValidation & "O"
      Case 16: sValidation = sValidation & "P"
      Case 17: sValidation = sValidation & "Q"
      Case 18: sValidation = sValidation & "R"
      Case 19: sValidation = sValidation & "S"
      Case 20: sValidation = sValidation & "T"
      Case 21: sValidation = sValidation & "U"
      Case 22: sValidation = sValidation & "V"
      Case 23: sValidation = sValidation & "W"
      Case 24: sValidation = sValidation & "X"
      Case 25: sValidation = sValidation & "Y"
      Case 26: sValidation = sValidation & "Z"
    End Select
  Next
  
  lblValidation.Caption = sValidation
End Sub

Private Sub cmdAnnuler_Click()
  m_frmSource.m_bValide = False
  
  Call Unload(Me)
End Sub

Private Sub cmdValider_Click()
  If UCase(lblValidation.Caption) = UCase(txtValidation.Text) Then
    m_frmSource.m_bValide = True
    
    Call Unload(Me)
  Else
    Call MsgBox("Le code de validation est incorrect!", vbOKOnly, "Erreur")
  End If
End Sub

Private Sub Form_Activate()
  Call txtValidation.SetFocus
End Sub

Private Sub Form_Load()
'  Call txtValidation.SetFocus
End Sub
'
