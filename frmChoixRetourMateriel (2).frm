VERSION 5.00
Begin VB.Form frmChoixRetourMateriel 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retour de mat�riel"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   3270
   Begin VB.CommandButton cmdFermer 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdMecanique 
      Caption         =   "M�canique"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdElectrique 
      Caption         =   "�lectrique"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   1080
      Width           =   1575
   End
End
Attribute VB_Name = "frmChoixRetourMateriel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdElectrique_Click()

 On Error GoTo Oups

 'Pour ouvrir le catalogue �lectrique
 Screen.MousePointer = vbHourglass

 Call frmRetourMateriel.Afficher(ELECTRIQUE)

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmChoixRetourMateriel", "cmdElectrique_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdMecanique_Click()

 On Error GoTo Oups

 'Pour ouvrir le catalogue m�canique
 Screen.MousePointer = vbHourglass

 Call frmRetourMateriel.Afficher(MECANIQUE)

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmChoixRetourMateriel", "cmdMecanique_Click", Err, Err.number, Err.Description
End Sub

Private Sub Cmdfermer_Click()

 On Error GoTo Oups

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmChoixRetourMateriel", "cmdFermer_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups

 Call Unload(frmChoixInventaire)

 Exit Sub

Oups:

 wOups "frmChoixRetourMateriel", "Form_Load", Err, Err.number, Err.Description
End Sub
