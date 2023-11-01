VERSION 5.00
Begin VB.Form frmAjoutDessin 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ajout d'un dessin"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4515
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmAjoutDessin.frx":0000
   ScaleHeight     =   2895
   ScaleWidth      =   4515
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAnnuler 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox txtDescription 
      Height          =   525
      Left            =   1920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox txtDessin 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Numéro du dessin :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   1695
   End
End
Attribute VB_Name = "frmAjoutDessin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
        
10      On Error GoTo AfficherErreur

20      If txtDessin.Text <> "" Then
30        frmDessins.m_sDessin = txtDessin.Text
40        frmDessins.m_sDescription = txtDescription.Text

50        frmDessins.m_bAnnuleAjout = False

60        Call Unload(Me)
70      Else
80        Call MsgBox("Le numéro est obligatoire!", vbOKOnly, "Erreur")
90      End If

100     Exit Sub

AfficherErreur:

110     Call AfficherErreur(Me, "cmdOK_Click", Err, Erl)
End Sub

Private Sub cmdAnnuler_Click()

10      On Error GoTo AfficherErreur

20      frmDessins.m_sDessin = ""
30      frmDessins.m_sDescription = ""

40      frmDessins.m_bAnnuleAjout = True

50      Call Unload(Me)

60      Exit Sub

AfficherErreur:

70      Call AfficherErreur(Me, "cmdAnnuler_Click", Err, Erl)
End Sub

Public Sub Afficher(ByVal sDessin As String, ByVal sDescription As String)
  txtDessin.Text = sDessin
  txtDescription.Text = sDescription
  
  Call Me.Show(vbModal)
End Sub
