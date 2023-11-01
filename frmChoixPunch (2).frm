VERSION 5.00
Begin VB.Form frmChoixPunch 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Punch"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   3270
   Begin VB.CommandButton cmdPunch 
      Caption         =   "Punch"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton cmdFeuilleTemps 
      Caption         =   "Feuilles de temps"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdFacturation 
      Caption         =   "Facturation"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton cmdFermer 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   3240
      Width           =   1575
   End
End
Attribute VB_Name = "frmChoixPunch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_sUserID As String 'Sert pour rechercher le userID de l'employé
Public m_iNoGroupe As Integer

Private Sub cmdFeuilleTemps_Click()

 On Error GoTo Oups

 'Ouverture du form pour l'impression des feuilles de temps
 Screen.MousePointer = vbHourglass
 
 Call OuvrirForm(frmFeuilleTemps, True)
 
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmChoixPunch", "cmdFeuilleTemps_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdFacturation_Click()

 On Error GoTo Oups

 'Ouverture du form pour la facturation des clients
 Screen.MousePointer = vbHourglass
 
 Call OuvrirForm(frmFacturation, True)
 
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmChoixPunch", "cmdFacturation_Click", Err, Err.number, Err.Description
End Sub

Private Sub Cmdfermer_Click()

 On Error GoTo Oups

 'Fermeture de la fenêtre
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmChoixPunch", "cmdFermer_Click", Err, Err.number, Err.Description
End Sub

Private Sub ActiverBoutonsGroupe()

 On Error GoTo Oups

 'Activation des boutons d'après le groupe
 'Bouton Punch
 cmdPunch.Enabled = g_bAffichagePunch
 
 'Bouton "Facturation"
 cmdFacturation.Enabled = g_bModificationFacturation

 Exit Sub

Oups:

 wOups "frmChoixPunch", "ActiverBoutonsGroupe", Err, Err.number, Err.Description
End Sub

Private Sub cmdPunch_Click()

 On Error GoTo Oups

 'Il faut afficher le login pour faire un punch in
 Call frmLogin.Afficher(Me)
 
 'Si bon password
 If g_bBonPasswd = True Then
 g_bBonPasswd = False
 
 'Ouverture du punch
 Call frmPunch.Afficher(m_sUserID)
 
 Call Unload(Me)
 End If

 Exit Sub

Oups:

 wOups "frmChoixPunch", "cmdPunch_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups

 'Ouverture de la fenêtre
 Call ActiverBoutonsGroupe

 Screen.MousePointer = vbDefault

 Exit Sub

Oups:

 wOups "frmChoixPunch", "Form_Load", Err, Err.number, Err.Description
End Sub
