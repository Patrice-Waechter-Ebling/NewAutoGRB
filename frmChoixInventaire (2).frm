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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   3255
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

 On Error GoTo Oups

 'Pour ouvrir l'inventaire électrique
 Screen.MousePointer = vbHourglass

 Call OuvrirForm(frmInventaireElec, False)
 
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmChoixInventaire", "cmdElectrique_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdMecanique_Click()

 On Error GoTo Oups

 'Pour ouvrir l'inventaire mécanique
 Screen.MousePointer = vbHourglass
 
 Call OuvrirForm(frmInventaireMec, False)
 
 Call Unload(Me)

 Exit Sub

Oups:

  wOups "frmChoixInventaire", "cmdMecanique_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdOutils_Click()

 On Error GoTo Oups

 'Inventaire des outils
 Screen.MousePointer = vbHourglass
 
 Call OuvrirForm(FrmOutils_InOut, False)
 
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmChoixInventaire", "cmdOutils_Click", Err, Err.number, Err.Description
End Sub

Private Sub Cmdfermer_Click()

 On Error GoTo Oups

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmChoixInventaire", "cmdFermer_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdRetour_Click()

 On Error GoTo Oups

 Call frmChoixRetourMateriel.Show(vbModal)

 Exit Sub

Oups:

 wOups "frmChoixInventaire", "cmdRetour_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdsortie_Click()

 On Error GoTo Oups

 Call frmChoixSortieMateriel.Show(vbModal)

 Exit Sub

Oups:

 wOups "frmChoixInventaire", "cmdSortie_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()
 
 On Error GoTo Oups

 Screen.MousePointer = vbDefault

 Exit Sub

Oups:

 wOups "frmChoixInventaire", "Form_Load", Err, Err.number, Err.Description
End Sub
