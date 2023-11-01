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
   MinButton       =   0   'False
   Picture         =   "frmChoixPunch.frx":0000
   ScaleHeight     =   3690
   ScaleWidth      =   3270
   StartUpPosition =   2  'CenterScreen
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

Public m_sUserID          As String 'Sert pour rechercher le userID de l'employé
Public m_iNoGroupe        As Integer

Private Sub cmdFeuilleTemps_Click()

5       On Error GoTo AfficherErreur

        'Ouverture du form pour l'impression des feuilles de temps
10      Screen.MousePointer = vbHourglass
  
15      Call OuvrirForm(frmFeuilleTemps, True)
  
20      Call Unload(Me)

25      Exit Sub

AfficherErreur:

30      woups "frmChoixPunch", "cmdFeuilleTemps_Click", Err, Erl
End Sub

Private Sub cmdFacturation_Click()

5       On Error GoTo AfficherErreur

        'Ouverture du form pour la facturation des clients
10      Screen.MousePointer = vbHourglass
  
15      Call OuvrirForm(frmFacturation, True)
  
20      Call Unload(Me)

25      Exit Sub

AfficherErreur:

30      woups "frmChoixPunch", "cmdFacturation_Click", Err, Erl
End Sub

Private Sub Cmdfermer_Click()

5       On Error GoTo AfficherErreur

        'Fermeture de la fenêtre
10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmChoixPunch", "cmdFermer_Click", Err, Erl
End Sub

Private Sub ActiverBoutonsGroupe()

5       On Error GoTo AfficherErreur

        'Activation des boutons d'après le groupe
        'Bouton Punch
10      cmdPunch.Enabled = g_bAffichagePunch
       
        'Bouton "Facturation"
15       cmdFacturation.Enabled = g_bModificationFacturation

20      Exit Sub

AfficherErreur:

25      woups "frmChoixPunch", "ActiverBoutonsGroupe", Err, Erl
End Sub

Private Sub cmdPunch_Click()

5       On Error GoTo AfficherErreur

        'Il faut afficher le login pour faire un punch in
10      Call frmLogin.Afficher(Me)
    
        'Si bon password
15      If g_bBonPasswd = True Then
20        g_bBonPasswd = False
   
          'Ouverture du punch
25        Call frmPunch.Afficher(m_sUserID)
    
30        Call Unload(Me)
35      End If

40      Exit Sub

AfficherErreur:

45      woups "frmChoixPunch", "cmdPunch_Click", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

        'Ouverture de la fenêtre
10      Call ActiverBoutonsGroupe

15      Screen.MousePointer = vbDefault

20      Exit Sub

AfficherErreur:

25      woups "frmChoixPunch", "Form_Load", Err, Erl
End Sub
