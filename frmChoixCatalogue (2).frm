VERSION 5.00
Begin VB.Form frmChoixCatalogue 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catalogue"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmChoixCatalogue.frx":0000
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
Attribute VB_Name = "frmChoixCatalogue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdElectrique_Click()

 On Error GoTo Oups

 'Pour ouvrir le catalogue électrique
 Screen.MousePointer = vbHourglass
 
 Call OuvrirForm(FrmCatalogueElec, False)
 
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmChoixCatalogue", "cmdElectrique_Click", Err, Err.number, Err.Description
End Sub

Private Sub cmdMecanique_Click()

 On Error GoTo Oups

 'Pour ouvrir le catalogue mécanique
 Screen.MousePointer = vbHourglass
 
 Call OuvrirForm(FrmCatalogueMec, False)
 
 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmChoixCatalogue", "cmdMecanique_Click", Err, Err.number, Err.Description
End Sub

Private Sub Cmdfermer_Click()

 On Error GoTo Oups

 Call Unload(Me)

 Exit Sub

Oups:

 wOups "frmChoixCatalogue", "cmdFermer_Click", Err, Err.number, Err.Description
End Sub

Private Sub Form_Load()

 On Error GoTo Oups

 Call ActiverBoutonsGroupe

 Screen.MousePointer = vbDefault

 Exit Sub

Oups:

 wOups "frmChoixCatalogue", "Form_Load", Err, Err.number, Err.Description
End Sub

Private Sub ActiverBoutonsGroupe()

 On Error GoTo Oups
 
 If g_bAffichageCatalogueMec = True Then
 cmdMecanique.Enabled = True
 Else
 cmdMecanique.Enabled = False
 End If
 
 If g_bAffichageCatalogueElec = True Then
 cmdElectrique.Enabled = True
 Else
 cmdElectrique.Enabled = False
 End If
 
  Exit Sub

Oups:

  wOups "frmChoixCatalogue", "ActiverBoutonsGroupe", Err, Err.number, Err.Description
End Sub
