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
   MDIChild        =   -1  'True
   ScaleHeight     =   4260
   ScaleWidth      =   4050
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

 On Error GoTo Oups

 If frmSource.Name = "FrmProjSoumElec" Then
 m_eType = ELECTRIQUE
 Else
 m_eType = MECANIQUE
 End If

 Call Me.Show(vbModal)

 Exit Sub

Oups:

 wOups "frmChoixDossier", "Afficher", Err, Err.number, Err.Description
End Sub

Private Sub cmdOK_Click()

 On Error GoTo Oups:

 If m_eType = ELECTRIQUE Then
 FrmProjSoumElec.m_bAnnulerChemin = False
 FrmProjSoumElec.m_sChemin = dirCheminPhotos.Path
 Else
 FrmProjSoumMec.m_bAnnulerChemin = False
 FrmProjSoumMec.m_sChemin = dirCheminPhotos.Path
 End If

 Call Unload(Me)
 
 Exit Sub
 
Oups:

 wOups "frmChoixDossier", "cmdOK_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnuler_Click()

 On Error GoTo Oups

 If m_eType = ELECTRIQUE Then
 FrmProjSoumElec.m_bAnnulerChemin = True
 FrmProjSoumElec.m_sChemin = vbNullString
 Else
 FrmProjSoumElec.m_bAnnulerChemin = True
 FrmProjSoumElec.m_sChemin = vbNullString
 End If

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmChoixDossier", "cmdAnnuler_Click", Err, Err.number, Err.Description
End Sub

Private Sub drvCheminPhotos_Change()

 On Error GoTo Oups

 dirCheminPhotos.Path = drvCheminPhotos.Drive

 Exit Sub

Oups:

 wOups "frmChoixDossier", "drvCheminPhotos_Change", Err, Err.number, Err.Description
End Sub
