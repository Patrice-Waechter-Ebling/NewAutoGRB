VERSION 5.00
Begin VB.Form frmChoixImpressionInventaire 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Quelle impression ?"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3870
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmChoixImpressionInventaire.frx":0000
   ScaleHeight     =   2160
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAnnuler 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdImprimer 
      Caption         =   "Imprimer"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
   Begin VB.OptionButton optValeurComptable 
      BackColor       =   &H00000000&
      Caption         =   "Valeurs comptables"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1935
   End
   Begin VB.OptionButton optAjustementInventaire 
      BackColor       =   &H00000000&
      Caption         =   "Ajustement de l'inventaire"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   2175
   End
End
Attribute VB_Name = "frmChoixImpressionInventaire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum enumImpressionInventaire
  MODE_AJUST_INV = 0
  MODE_VAL_COMPTABLE = 1
End Enum

Private m_frmSource As Form

Private Sub cmdAnnuler_Click()

5       On Error GoTo AfficherErreur

10      m_frmSource.m_bAnnuleImpression = True

15      Call Unload(Me)

20      Exit Sub

AfficherErreur:

25      woups "frmChoixImpressionInventaire", "cmdAnnuler_Click", Err, Erl
End Sub

Private Sub cmdImprimer_Click()
  
5       On Error GoTo AfficherErreur
  
10      m_frmSource.m_bAnnuleImpression = False

15      If optValeurComptable.Value = True Then
20        m_frmSource.m_eTypeImpression = MODE_VAL_COMPTABLE
25      Else
30        m_frmSource.m_eTypeImpression = MODE_AJUST_INV
35      End If
    
40      Call Unload(Me)

45      Exit Sub

AfficherErreur:

50      woups "frmChoixImpressionInventaire", "cmdImprimer_Click", Err, Erl
End Sub

Public Sub Afficher(ByVal frmSource As Form)

5       On Error GoTo AfficherErreur

10      Set m_frmSource = frmSource

15      Call Me.Show(vbModal)

20      Exit Sub

AfficherErreur:

25      woups "frmChoixImpressionInventaire", "Afficher", Err, Erl
End Sub

Private Sub Form_Load()
    If m_frmSource.m_typeImpressionExel Then
       cmdImprimer.Caption = "Exporter"
    Else
        cmdImprimer.Caption = "Imprimer"
    End If

End Sub
