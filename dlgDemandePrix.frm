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

5       On Error GoTo AfficherErreur

10      Set m_objForm = objForm
    
15      Call Show(vbModal)

20      Exit Sub

AfficherErreur:

25      woups "dlgDemandePrix", "Afficher", Err, Erl
End Sub

Private Sub cmdNouveau_Click()

5       On Error GoTo AfficherErreur

10      m_objForm.m_eDemande = MODE_NOUVELLE
  
15      m_objForm.m_bDemandeAnnuler = False
  
20      Call Unload(Me)

25      Exit Sub

AfficherErreur:

30      woups "dlgDemandePrix", "cmdNouveau_Click", Err, Erl
End Sub

Private Sub cmdPiece_Click()

5       On Error GoTo AfficherErreur
       
10      m_objForm.m_eDemande = MODE_PIECE
  
15      m_objForm.m_bDemandeAnnuler = False
  
20      Call Unload(Me)

25      Exit Sub

AfficherErreur:

30      woups "dlgDemandePrix", "cmdPiece_Click", Err, Erl
End Sub

Private Sub cmdFournisseur_Click()

5       On Error GoTo AfficherErreur
       
10      m_objForm.m_eDemande = MODE_FOURNISSEUR
  
15      m_objForm.m_bDemandeAnnuler = False
  
20      Call Unload(Me)

25      Exit Sub

AfficherErreur:

30      woups "dlgDemandePrix", "cmdFournisseur_Click", Err, Erl
End Sub

Private Sub cmdCategorie_Click()

5       On Error GoTo AfficherErreur

10      m_objForm.m_eDemande = MODE_CATEGORIE
  
15      m_objForm.m_bDemandeAnnuler = False
  
20      Call Unload(Me)

25      Exit Sub

AfficherErreur:

30      woups "dlgDemandePrix", "cmdCategorie_Click", Err, Erl
End Sub

Private Sub cmdAnnuler_Click()

5       On Error GoTo AfficherErreur

10      m_objForm.m_bDemandeAnnuler = True
  
15      Call Unload(Me)

20      Exit Sub

AfficherErreur:

25     woups "dlgDemandePrix", "cmdAnnuler_Click", Err, Erl
End Sub
