VERSION 5.00
Begin VB.Form frmChoixImpressionRetourMarchandise 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choix d'impression"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmChoixImpressionRetourMarchandise.frx":0000
   ScaleHeight     =   2160
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optDemande 
      BackColor       =   &H00000000&
      Caption         =   "Demande de retour de marchandise"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2895
   End
   Begin VB.OptionButton optRetour 
      BackColor       =   &H00000000&
      Caption         =   "Retour de marchandise"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   2895
   End
   Begin VB.CommandButton cmdImprimer 
      Caption         =   "Imprimer"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdAnnuler 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   1440
      Width           =   975
   End
End
Attribute VB_Name = "frmChoixImpressionRetourMarchandise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum enumImpressionRetour
  MODE_DEMANDE_RETOUR = 0
  MODE_RETOUR = 1
End Enum

Private Sub cmdAnnuler_Click()

5       On Error GoTo AfficherErreur

10      frmRetourMarchandise.m_bAnnuleImpression = True

15      Call Unload(Me)

20      Exit Sub

AfficherErreur:

25      woups "frmChoixImpressionRetourMarchandise", "cmdAnnuler_Click", Err, Erl
End Sub

Private Sub cmdImprimer_Click()

5       On Error GoTo AfficherErreur

10      frmRetourMarchandise.m_bAnnuleImpression = False

15      If optRetour.Value = True Then
20        frmRetourMarchandise.m_eTypeImpression = MODE_RETOUR
25      Else
30        frmRetourMarchandise.m_eTypeImpression = MODE_DEMANDE_RETOUR
35      End If
    
40      Call Unload(Me)

45      Exit Sub

AfficherErreur:

50      woups "frmChoixImpressionRetourMarchandise", "cmdImprimer_Click", Err, Erl
End Sub

