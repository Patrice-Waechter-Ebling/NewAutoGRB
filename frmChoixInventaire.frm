VERSION 5.00
Begin VB.Form frmChoixInventaire 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventaire"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3255
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmChoixInventaire.frx":0000
   ScaleHeight     =   3750
   ScaleWidth      =   3255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFermer 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton cmdOutils 
      Caption         =   "Outils"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton cmdElectrique 
      Caption         =   "Électrique"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton cmdMecanique 
      Caption         =   "Mécanique"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton cmdSortie 
      Caption         =   "Sortie de matériel"
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   2520
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdRetour 
      Caption         =   "Retour de matériel"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   3000
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "frmChoixInventaire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdElectrique_Click()

5       On Error GoTo AfficherErreur

        'Pour ouvrir l'inventaire électrique
10      Screen.MousePointer = vbHourglass

15      Call OuvrirForm(frmInventaireElec, False)
   
20      Call Unload(Me)

25      Exit Sub

AfficherErreur:

30      woups "frmChoixInventaire", "cmdElectrique_Click", Err, Erl
End Sub

Private Sub cmdMecanique_Click()

5       On Error GoTo AfficherErreur

        'Pour ouvrir l'inventaire mécanique
10      Screen.MousePointer = vbHourglass
  
15      Call OuvrirForm(frmInventaireMec, False)
  
20      Call Unload(Me)

25      Exit Sub

AfficherErreur:

60      woups "frmChoixInventaire", "cmdMecanique_Click", Err, Erl
End Sub

Private Sub cmdOutils_Click()

5       On Error GoTo AfficherErreur

        'Inventaire des outils
10      Screen.MousePointer = vbHourglass
  
15      Call OuvrirForm(FrmOutils_InOut, False)
  
20      Call Unload(Me)

25      Exit Sub

AfficherErreur:

30      woups "frmChoixInventaire", "cmdOutils_Click", Err, Erl
End Sub

Private Sub Cmdfermer_Click()

5       On Error GoTo AfficherErreur

10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmChoixInventaire", "cmdFermer_Click", Err, Erl
End Sub

Private Sub cmdRetour_Click()

5       On Error GoTo AfficherErreur

10      Call frmChoixRetourMateriel.Show(vbModal)

15      Exit Sub

AfficherErreur:

20      woups "frmChoixInventaire", "cmdRetour_Click", Err, Erl
End Sub

Private Sub cmdsortie_Click()

5       On Error GoTo AfficherErreur

10      Call frmChoixSortieMateriel.Show(vbModal)

15      Exit Sub

AfficherErreur:

20      woups "frmChoixInventaire", "cmdSortie_Click", Err, Erl
End Sub

Private Sub Form_Load()
  
5       On Error GoTo AfficherErreur

10      Screen.MousePointer = vbDefault

15      Exit Sub

AfficherErreur:

20      woups "frmChoixInventaire", "Form_Load", Err, Erl
End Sub
