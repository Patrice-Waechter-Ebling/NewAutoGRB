VERSION 5.00
Begin VB.Form frmChoixDossier 
   BackColor       =   &H00000000&
   Caption         =   "Choix du dossier"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4050
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmChoixDossier.frx":0000
   ScaleHeight     =   4260
   ScaleWidth      =   4050
   StartUpPosition =   2  'CenterScreen
   Begin VB.DriveListBox drvCheminPhotos 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   3615
   End
   Begin VB.CommandButton cmdAnnuler 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   3840
      Width           =   975
   End
   Begin VB.DirListBox dirCheminPhotos 
      Height          =   2565
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   3615
   End
End
Attribute VB_Name = "frmChoixDossier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enumType
  ELECTRIQUE = 0
  MECANIQUE = 1
End Enum

Private m_eType As enumType

Public Sub Afficher(ByVal frmSource As Form)

5       On Error GoTo AfficherErreur

10      If frmSource.Name = "FrmProjSoumElec" Then
15        m_eType = ELECTRIQUE
20      Else
25        m_eType = MECANIQUE
30      End If

35      Call Me.Show(vbModal)

40      Exit Sub

AfficherErreur:

45      woups "frmChoixDossier", "Afficher", Err, Erl
End Sub

Private Sub cmdOK_Click()

5       On Error GoTo AfficherErreur:

10      If m_eType = ELECTRIQUE Then
15        FrmProjSoumElec.m_bAnnulerChemin = False
20        FrmProjSoumElec.m_sChemin = dirCheminPhotos.Path
25      Else
30        FrmProjSoumMec.m_bAnnulerChemin = False
35        FrmProjSoumMec.m_sChemin = dirCheminPhotos.Path
40      End If

45      Call Unload(Me)
  
50      Exit Sub
  
AfficherErreur:

55      woups "frmChoixDossier", "cmdOK_Click", Err, Erl
End Sub

Private Sub cmdAnnuler_Click()

5       On Error GoTo AfficherErreur

10      If m_eType = ELECTRIQUE Then
15        FrmProjSoumElec.m_bAnnulerChemin = True
20        FrmProjSoumElec.m_sChemin = vbNullString
25      Else
30        FrmProjSoumElec.m_bAnnulerChemin = True
35        FrmProjSoumElec.m_sChemin = vbNullString
40      End If

45      Call Unload(Me)

50      Exit Sub

AfficherErreur:

55      woups "frmChoixDossier", "cmdAnnuler_Click", Err, Erl
End Sub

Private Sub drvCheminPhotos_Change()

5       On Error GoTo AfficherErreur

10      dirCheminPhotos.Path = drvCheminPhotos.Drive

15      Exit Sub

AfficherErreur:

20      woups "frmChoixDossier", "drvCheminPhotos_Change", Err, Erl
End Sub
