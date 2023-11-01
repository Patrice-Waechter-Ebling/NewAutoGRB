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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
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

 On Error GoTo Oups

 frmRetourMarchandise.m_bAnnuleImpression = True

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmChoixImpressionRetourMarchandise", "cmdAnnuler_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdImprimer_Click()

 On Error GoTo Oups

 frmRetourMarchandise.m_bAnnuleImpression = False

 If optRetour.Value = True Then
 frmRetourMarchandise.m_eTypeImpression = MODE_RETOUR
 Else
 frmRetourMarchandise.m_eTypeImpression = MODE_DEMANDE_RETOUR
 End If
 
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmChoixImpressionRetourMarchandise", "cmdImprimer_Click", Err, Err.number, Err.Description
End Sub

