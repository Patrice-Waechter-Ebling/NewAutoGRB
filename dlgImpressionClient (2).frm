VERSION 5.00
Begin VB.Form dlgImpressionClient 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impression des clients"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8370
   Icon            =   "dlgImpressionClient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   8370
   Begin VB.CommandButton cmdFacturer 
      Caption         =   "Clients Facturés"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdPotentiel 
      Caption         =   "Clients Potentiels"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdCategorie 
      Caption         =   "Catégorie"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdAnnuler 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   6960
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdTous 
      Caption         =   "Tous les clients"
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdVille 
      Caption         =   "Ville"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Quel est le tri de l'impression des clients ?"
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
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7935
   End
End
Attribute VB_Name = "dlgImpressionClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFacturer_Click()

 On Error GoTo Oups
 
 FrmClient.m_bImpressionVille = False
 FrmClient.m_bImpressionAnnuler = False
 FrmClient.m_bImpressionCategorie = False
 FrmClient.m_bImpressionPotentiel = False
 FrmClient.m_bImpressionFacturer = True
 
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "dlgImpressionClient", "cmdVille_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdPotentiel_Click()

 On Error GoTo Oups
 
 FrmClient.m_bImpressionVille = False
 FrmClient.m_bImpressionAnnuler = False
 FrmClient.m_bImpressionCategorie = False
 FrmClient.m_bImpressionPotentiel = True
 FrmClient.m_bImpressionFacturer = False
 
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "dlgImpressionClient", "cmdVille_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdVille_Click()

 On Error GoTo Oups
 
 FrmClient.m_bImpressionVille = True
 FrmClient.m_bImpressionAnnuler = False
 FrmClient.m_bImpressionCategorie = False
 FrmClient.m_bImpressionPotentiel = False
 FrmClient.m_bImpressionFacturer = False
 
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "dlgImpressionClient", "cmdVille_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdTous_Click()

 On Error GoTo Oups
 
 FrmClient.m_bImpressionAnnuler = False
 FrmClient.m_bImpressionVille = False
 FrmClient.m_bImpressionCategorie = False
 FrmClient.m_bImpressionPotentiel = False
 FrmClient.m_bImpressionFacturer = False
 
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "dlgImpressionClient", "cmdTous_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdCategorie_Click()

 On Error GoTo Oups

 FrmClient.m_bImpressionCategorie = True
 FrmClient.m_bImpressionAnnuler = False
 FrmClient.m_bImpressionVille = False
 FrmClient.m_bImpressionPotentiel = False
 FrmClient.m_bImpressionFacturer = False
 
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "dlgImpressionClient", "cmdCategorie_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnuler_Click()

 On Error GoTo Oups

 FrmClient.m_bImpressionAnnuler = True
 FrmClient.m_bImpressionCategorie = False
 FrmClient.m_bImpressionVille = False
 FrmClient.m_bImpressionPotentiel = False
 FrmClient.m_bImpressionFacturer = False
 
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "dlgImpressionClient", "cmdAnnuler_Click", Err, Err.number, Err.Description
End Sub
