VERSION 5.00
Begin VB.Form frmChoixRetourMateriel 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retour de matériel"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmChoixRetourMateriel.frx":0000
   ScaleHeight     =   3165
   ScaleWidth      =   3270
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFermer 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdMecanique 
      Caption         =   "Mécanique"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdElectrique 
      Caption         =   "Électrique"
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

5       On Error GoTo AfficherErreur

        'Pour ouvrir le catalogue électrique
10      Screen.MousePointer = vbHourglass

15      Call frmRetourMateriel.Afficher(ELECTRIQUE)

20      Call Unload(Me)

25      Exit Sub

AfficherErreur:

30      woups "frmChoixRetourMateriel", "cmdElectrique_Click", Err, Erl
End Sub

Private Sub cmdMecanique_Click()

5       On Error GoTo AfficherErreur

        'Pour ouvrir le catalogue mécanique
10      Screen.MousePointer = vbHourglass

15      Call frmRetourMateriel.Afficher(MECANIQUE)

20      Call Unload(Me)

25      Exit Sub

AfficherErreur:

30      woups "frmChoixRetourMateriel", "cmdMecanique_Click", Err, Erl
End Sub

Private Sub Cmdfermer_Click()

5       On Error GoTo AfficherErreur

10      Call Unload(Me)

15      Exit Sub

AfficherErreur:

20      woups "frmChoixRetourMateriel", "cmdFermer_Click", Err, Erl
End Sub

Private Sub Form_Load()

5       On Error GoTo AfficherErreur

10      Call Unload(frmChoixInventaire)

15      Exit Sub

AfficherErreur:

20      woups "frmChoixRetourMateriel", "Form_Load", Err, Erl
End Sub
