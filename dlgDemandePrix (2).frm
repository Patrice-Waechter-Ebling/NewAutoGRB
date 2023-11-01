VERSION 5.00
Begin VB.Form dlgDemandePrix 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demande de prix"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   7950
   Begin VB.CommandButton cmdNouveau 
      Caption         =   "Nouvelles pièces"
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdCategorie 
      Caption         =   "Une catégorie"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdAnnuler 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   6360
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdFournisseur 
      Caption         =   "Toutes les pièces"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdPiece 
      Caption         =   "Une pièce"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Voulez-vous faire une demande de prix pour tous les pièces d'un fournisseur, d'une catégorie ou pour une pièce en particulier?"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "dlgDemandePrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_objForm As Form

Public Sub Afficher(ByVal objForm As Form)

 On Error GoTo Oups

 Set m_objForm = objForm
 
 Call Show(vbModal)

 Exit Sub

Oups:

 wOups "dlgDemandePrix", "Afficher", Err, Err.number, Err.Description
End Sub

Private Sub cmdNouveau_Click()

 On Error GoTo Oups

 m_objForm.m_eDemande = MODE_NOUVELLE
 
 m_objForm.m_bDemandeAnnuler = False
 
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "dlgDemandePrix", "cmdNouveau_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdPiece_Click()

 On Error GoTo Oups
 
 m_objForm.m_eDemande = MODE_PIECE
 
 m_objForm.m_bDemandeAnnuler = False
 
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "dlgDemandePrix", "cmdPiece_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdFournisseur_Click()

 On Error GoTo Oups
 
 m_objForm.m_eDemande = MODE_FOURNISSEUR
 
 m_objForm.m_bDemandeAnnuler = False
 
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "dlgDemandePrix", "cmdFournisseur_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdCategorie_Click()

 On Error GoTo Oups

 m_objForm.m_eDemande = MODE_CATEGORIE
 
 m_objForm.m_bDemandeAnnuler = False
 
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "dlgDemandePrix", "cmdCategorie_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdAnnuler_Click()

 On Error GoTo Oups

 m_objForm.m_bDemandeAnnuler = True
 
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "dlgDemandePrix", "cmdAnnuler_Click", Err, Err.number, Err.Description
End Sub
